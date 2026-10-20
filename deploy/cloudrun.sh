#!/usr/bin/env bash
# Deploy the remote tracker-mcp server to Google Cloud Run. Run it in Cloud Shell
# (shell.cloud.google.com), from a clone of the repository:
#
#   git clone https://github.com/butyavk/tracker-mcp && cd tracker-mcp
#   ./deploy/cloudrun.sh --project <project-id> --client-id <web client id> \
#       --allow you@gmail.com,friend@gmail.com [--picker-key <API key>] [--region europe-west1]
#
# Before: billing enabled on the project; an OAuth client of type "Web application" with
# the redirect URI this script prints (run it once with --dry-run to see it). The client
# secret is asked for here, without echo, and kept in Secret Manager; it is never printed.
set -euo pipefail

REGION=europe-west1 SERVICE=tracker-mcp PICKER_KEY="" DRY=0 PROJECT="" CLIENT_ID="" ALLOW=""
while [ $# -gt 0 ]; do
  case "$1" in
    --project) PROJECT=$2; shift 2;;
    --region) REGION=$2; shift 2;;
    --client-id) CLIENT_ID=$2; shift 2;;
    --allow) ALLOW=$2; shift 2;;
    --picker-key) PICKER_KEY=$2; shift 2;;
    --service) SERVICE=$2; shift 2;;
    --dry-run) DRY=1; shift;;
    *) echo "unknown option $1" >&2; exit 2;;
  esac
done
[ -n "$PROJECT" ] || PROJECT=$(gcloud config get-value project 2>/dev/null)
[ -n "$PROJECT" ] || { echo "pass --project" >&2; exit 2; }
NUMBER=$(gcloud projects describe "$PROJECT" --format='value(projectNumber)')
BASE="https://${SERVICE}-${NUMBER}.${REGION}.run.app"
echo "Server address:        $BASE"
echo "Connector URL:         $BASE/mcp"
echo "OAuth redirect URI:    $BASE/google/callback   (Web application client)"
echo "JavaScript origin:     $BASE   (for the Drive file picker)"
[ "$DRY" = 1 ] && exit 0
[ -n "$CLIENT_ID" ] && [ -n "$ALLOW" ] || { echo "pass --client-id and --allow" >&2; exit 2; }

gcloud config set project "$PROJECT" >/dev/null
echo "Enabling APIs (a minute the first time)…"
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com \
  secretmanager.googleapis.com storage.googleapis.com sheets.googleapis.com picker.googleapis.com >/dev/null

# source deploys build with the Compute Engine default service account (Cloud Build's
# default since 2024); it needs the Cloud Run Builder role
gcloud projects add-iam-policy-binding "$PROJECT" --condition=None \
  --member "serviceAccount:${NUMBER}-compute@developer.gserviceaccount.com" --role roles/run.builder >/dev/null

BUCKET="${PROJECT}-tracker-mcp-data"
if ! gcloud storage buckets describe "gs://$BUCKET" >/dev/null 2>&1; then
  gcloud storage buckets create "gs://$BUCKET" --location="$REGION" --uniform-bucket-level-access \
    --public-access-prevention >/dev/null
  echo "Bucket gs://$BUCKET created (private)."
fi

SA="tracker-mcp-run@${PROJECT}.iam.gserviceaccount.com"
gcloud iam service-accounts describe "$SA" >/dev/null 2>&1 || \
  gcloud iam service-accounts create tracker-mcp-run --display-name "tracker-mcp remote server" >/dev/null
gcloud storage buckets add-iam-policy-binding "gs://$BUCKET" --member "serviceAccount:$SA" \
  --role roles/storage.objectUser >/dev/null

secret() {   # name, value from stdin; only creates or adds a version
  if gcloud secrets describe "$1" >/dev/null 2>&1; then
    gcloud secrets versions add "$1" --data-file=- >/dev/null
  else
    gcloud secrets create "$1" --replication-policy=automatic --data-file=- >/dev/null
  fi
  gcloud secrets add-iam-policy-binding "$1" --member "serviceAccount:$SA" \
    --role roles/secretmanager.secretAccessor >/dev/null
}
if ! gcloud secrets describe tracker-mcp-secret-key >/dev/null 2>&1; then
  openssl rand -base64 32 | tr '+/' '-_' | tr -d '\n' | secret tracker-mcp-secret-key
  echo "Encryption key generated and stored (tracker-mcp-secret-key)."
fi
if ! gcloud secrets describe tracker-mcp-google-client-secret >/dev/null 2>&1; then
  read -rsp "Paste the Web client's secret (not shown): " CS; echo
  printf '%s' "$CS" | secret tracker-mcp-google-client-secret; unset CS
fi

# "^|^" makes | the separator, so the e-mail list keeps its commas
ENV="^|^TRACKER_MCP_BASE_URL=$BASE|TRACKER_MCP_DATA_DIR=/data|GOOGLE_WEB_CLIENT_ID=$CLIENT_ID|GOOGLE_APP_ID=$NUMBER"
ENV="$ENV|TRACKER_MCP_ALLOWED_EMAILS=$ALLOW"
[ -n "$PICKER_KEY" ] && ENV="$ENV|GOOGLE_PICKER_API_KEY=$PICKER_KEY"

echo "Building and deploying (3–5 minutes)…"
gcloud run deploy "$SERVICE" --source . --region "$REGION" --service-account "$SA" \
  --allow-unauthenticated --max-instances 1 --min-instances 0 --memory 512Mi --cpu 1 \
  --execution-environment gen2 --timeout 300 --quiet \
  --add-volume name=data,type=cloud-storage,bucket="$BUCKET" --add-volume-mount volume=data,mount-path=/data \
  --set-env-vars "$ENV" \
  --set-secrets TRACKER_MCP_SECRET_KEY=tracker-mcp-secret-key:latest,GOOGLE_WEB_CLIENT_SECRET=tracker-mcp-google-client-secret:latest

URL=$(gcloud run services describe "$SERVICE" --region "$REGION" --format='value(status.url)')
echo
echo "Done. Health check: $(curl -s "$BASE/healthz")"
[ "$URL" = "$BASE" ] || echo "Note: Cloud Run reports $URL; the server is configured for $BASE."
echo "In Claude: Settings → Connectors → Add custom connector → $BASE/mcp"
