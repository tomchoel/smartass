// A slow headless "claude" for the pane server test (CLAUDE_JS): one stream-json line, then the run stays open for
// FAKE_CLAUDE_MS and exits 0. No AI, no tools.
process.stdout.write(JSON.stringify({ type: "system", subtype: "init", session_id: "fake" }) + "\n")
setTimeout(() => process.exit(0), Number(process.env.FAKE_CLAUDE_MS || 1500))
