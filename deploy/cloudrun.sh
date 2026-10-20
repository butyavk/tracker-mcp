#!/usr/bin/env bash
# Deploy the remote tracker-mcp server to Google Cloud Run. Run it in Cloud Shell
# (shell.cloud.google.com), from a clone of the repository:
#
#   git clone https://github.com/butyavk/tracker-mcp && cd tracker-mcp
#   ./deploy/cloudrun.sh --project <project-id> --client-id <web client id> \
#       --allow you@gmail.com,friend@gmail.com [--region europe-west1]
#
# Before: billing enabled on the project; an OAuth client of type "Web application" with
# the redirect URI this script prints (run it once with --dry-run to see it). The client
# secret is asked for here, without echo, and kept in Secret Manager; it is never printed.
# The file picker's API key is made here too (or pass your own with --picker-key).
set -euo pipefail

REGION=europe-west1 SERVICE=tracker-mcp PICKER_KEY="" DRY=0 PROJECT="" CLIENT_ID="" ALLOW="" NEW_SECRET=0
while [ $# -gt 0 ]; do
  case "$1" in
    --project) PROJECT=$2; shift 2;;
    --region) REGION=$2; shift 2;;
    --client-id) CLIENT_ID=$2; shift 2;;
    --allow) ALLOW=$2; shift 2;;
    --picker-key) PICKER_KEY=$2; shift 2;;
    --service) SERVICE=$2; shift 2;;
    --dry-run) DRY=1; shift;;
    --new-client-secret) NEW_SECRET=1; shift;;
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

export CLOUDSDK_CORE_PROJECT="$PROJECT"   # for this script only; your gcloud setting stays as it was
echo "Enabling APIs (a minute the first time)…"
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com \
  secretmanager.googleapis.com storage.googleapis.com iam.googleapis.com compute.googleapis.com \
  sheets.googleapis.com picker.googleapis.com >/dev/null   # compute: its default account runs the build

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
if ! gcloud iam service-accounts describe "$SA" >/dev/null 2>&1; then
  gcloud iam service-accounts create tracker-mcp-run --display-name "tracker-mcp remote server" >/dev/null
  sleep 15   # a new service account takes a moment to show up for IAM bindings
fi
gcloud storage buckets add-iam-policy-binding "gs://$BUCKET" --member "serviceAccount:$SA" \
  --role roles/storage.objectUser >/dev/null

secret() {   # name, value from stdin; creates the secret or adds a version
  if gcloud secrets describe "$1" >/dev/null 2>&1; then
    gcloud secrets versions add "$1" --data-file=- >/dev/null
  else
    gcloud secrets create "$1" --replication-policy=automatic --data-file=- >/dev/null
  fi
}
if ! gcloud secrets describe tracker-mcp-secret-key >/dev/null 2>&1; then
  openssl rand -base64 32 | tr '+/' '-_' | tr -d '\n' | secret tracker-mcp-secret-key
  echo "Encryption key generated and stored (tracker-mcp-secret-key)."
fi
if [ "$NEW_SECRET" = 1 ] || ! gcloud secrets describe tracker-mcp-google-client-secret >/dev/null 2>&1; then
  read -rsp "Paste the Web client's secret (not shown): " CS; echo
  printf '%s' "$CS" | secret tracker-mcp-google-client-secret; unset CS
fi
for s in tracker-mcp-secret-key tracker-mcp-google-client-secret; do
  gcloud secrets add-iam-policy-binding "$s" --member "serviceAccount:$SA" \
    --role roles/secretmanager.secretAccessor >/dev/null
done

# The Drive file picker needs a browser API key. Without --picker-key the script makes one
# (or reuses it), limited to the Picker API and to this server's pages. It only names the
# app to Google; every browser that opens the picker sees it anyway.
if [ -z "$PICKER_KEY" ]; then
  gcloud services enable apikeys.googleapis.com >/dev/null
  key_name() { gcloud services api-keys list --filter='displayName="tracker-mcp picker"' \
    --format='value(name)' --limit=1; }
  KEY=$(key_name)
  if [ -z "$KEY" ]; then
    out=$(gcloud services api-keys create --display-name="tracker-mcp picker" \
      --api-target=service=picker.googleapis.com --allowed-referrers="$BASE/*" 2>&1) \
      || { echo "$out" >&2; exit 1; }
    KEY=$(key_name)
    echo "Picker API key created (tracker-mcp picker: Picker API only, $BASE only)."
  fi
  PICKER_KEY=$(gcloud services api-keys get-key-string "$KEY" --format='value(keyString)')
fi

# "^|^" makes | the separator, so the e-mail list keeps its commas
ENV="^|^TRACKER_MCP_BASE_URL=$BASE|TRACKER_MCP_DATA_DIR=/data|GOOGLE_WEB_CLIENT_ID=$CLIENT_ID|GOOGLE_APP_ID=$NUMBER"
ENV="$ENV|TRACKER_MCP_ALLOWED_EMAILS=$ALLOW"
ENV="$ENV|GOOGLE_PICKER_API_KEY=$PICKER_KEY"

echo "Building and deploying (3–5 minutes)…"
gcloud run deploy "$SERVICE" --source . --region "$REGION" --service-account "$SA" \
  --allow-unauthenticated --max-instances 1 --min-instances 0 --memory 512Mi --cpu 1 \
  --execution-environment gen2 --timeout 300 --quiet \
  --add-volume "name=data,type=cloud-storage,bucket=$BUCKET,mount-options=uid=10001;gid=10001" \
  --add-volume-mount volume=data,mount-path=/data \
  --set-env-vars "$ENV" \
  --set-secrets TRACKER_MCP_SECRET_KEY=tracker-mcp-secret-key:latest,GOOGLE_WEB_CLIENT_SECRET=tracker-mcp-google-client-secret:latest

echo
HEALTH=$(curl -s --max-time 20 "$BASE/health" || true)
echo "Done. Health check: ${HEALTH:-no answer}"
if [ "$HEALTH" != ok ]; then   # Cloud Run also serves a hashed address; the server only answers as $BASE
  URL=$(gcloud run services describe "$SERVICE" --region "$REGION" --format='value(status.url)')
  echo "Note: $BASE did not answer yet; Cloud Run reports $URL. Try the health check again in a minute."
fi
echo "In Claude: Settings → Connectors → Add custom connector → $BASE/mcp"
