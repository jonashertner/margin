#!/bin/zsh
cd "${0:A:h}"
NODE_BIN="$(command -v node)"
if [[ -z "$NODE_BIN" && -x /opt/homebrew/bin/node ]]; then
  NODE_BIN=/opt/homebrew/bin/node
fi
if [[ -z "$NODE_BIN" ]]; then
  print 'Margin needs Node.js. See README.md.'
  read '?Press Return to close.'
  exit 1
fi
print 'Margin · Legal revision'
print 'Open http://127.0.0.1:4317 in your browser.'
print 'Keep this window open while reviewing. Press Control-C to stop.'
"$NODE_BIN" server.mjs
