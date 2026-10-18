/*
 * Launcher for running the tracker MCP server from Claude Desktop on macOS.
 *
 * Why it exists: Claude Desktop starts a local MCP server as its own process,
 * so the server does not inherit Claude's privacy permissions. macOS then
 * treats a bare /bin/sh or python as the "responsible" process. A bare binary
 * cannot ask for access to ~/Documents, so the request is denied silently
 * ("Operation not permitted"), and the tracker lives in ~/Documents.
 *
 * This launcher is the main executable of a tiny app bundle
 * ("Tracker MCP.app"). It starts the real server as a child process, and
 * children inherit the responsible process, so macOS asks once, on behalf of
 * "Tracker MCP": "wants to access files in your Documents folder". Only this
 * app gets the permission, not Claude as a whole.
 *
 * The launcher does not copy any data: the child inherits stdin/stdout (the
 * MCP channel) directly. It forwards termination signals to the child and
 * exits with the child's status. Diagnostics go to stderr, which Claude
 * Desktop writes to ~/Library/Logs/Claude/mcp-server-<name>.log.
 *
 * Usage: tracker-mcp /bin/sh /path/to/run.sh
 */
#define _GNU_SOURCE
#include <dlfcn.h>
#include <errno.h>
#include <signal.h>
#include <spawn.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/wait.h>
#include <unistd.h>

extern char **environ;
static pid_t child = 0;

static void forward(int sig) {
    if (child > 0) kill(child, sig);
}

static void diagnostics(void) {
#ifdef __APPLE__
    /* sandbox_check(pid, NULL, 0) != 0 means the process runs in a seatbelt
       sandbox; then no privacy permission can help and this approach fails. */
    int (*sandbox_check)(pid_t, const char *, int, ...) =
        (int (*)(pid_t, const char *, int, ...))dlsym(RTLD_DEFAULT, "sandbox_check");
    if (sandbox_check)
        fprintf(stderr, "tracker-mcp launcher: pid %d, sandboxed=%d\n",
                (int)getpid(), sandbox_check(getpid(), NULL, 0));
#endif
    /* Touching the file here also triggers the macOS permission prompt early. */
    const char *path = getenv("TRACKER_PATH");
    if (path)
        fprintf(stderr, "tracker-mcp launcher: tracker readable: %s\n",
                access(path, R_OK) == 0 ? "yes" : strerror(errno));
}

int main(int argc, char *argv[]) {
    if (argc < 2) {
        fprintf(stderr, "usage: %s /absolute/command [args...]\n", argv[0]);
        return 2;
    }
    diagnostics();

    struct sigaction sa;
    memset(&sa, 0, sizeof sa);
    sa.sa_handler = forward;
    sigaction(SIGTERM, &sa, NULL);
    sigaction(SIGINT, &sa, NULL);
    sigaction(SIGHUP, &sa, NULL);

    int rc = posix_spawn(&child, argv[1], NULL, NULL, &argv[1], environ);
    if (rc != 0) {
        fprintf(stderr, "tracker-mcp launcher: cannot start %s: %s\n", argv[1], strerror(rc));
        return 127;
    }
    int status;
    while (waitpid(child, &status, 0) < 0) {
        if (errno != EINTR) return 1;
    }
    if (WIFEXITED(status)) return WEXITSTATUS(status);
    if (WIFSIGNALED(status)) return 128 + WTERMSIG(status);
    return 1;
}
