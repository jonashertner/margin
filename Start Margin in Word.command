#!/bin/zsh

if ! cd "${0:A:h}"; then
  print 'Margin could not open its application folder.'
  exit 1
fi

margin_pause() {
  if [[ -t 0 ]]; then
    read '?Press Return to close.'
  fi
}

MARGIN_NODE_BIN="$(command -v node)"
if [[ -z "$MARGIN_NODE_BIN" && -x /opt/homebrew/bin/node ]]; then
  MARGIN_NODE_BIN=/opt/homebrew/bin/node
elif [[ -z "$MARGIN_NODE_BIN" && -x /usr/local/bin/node ]]; then
  MARGIN_NODE_BIN=/usr/local/bin/node
fi
if [[ -z "$MARGIN_NODE_BIN" ]]; then
  print 'Margin needs Node.js 22 or later. See README.md, then open this launcher again.'
  margin_pause
  exit 1
fi
if ! "$MARGIN_NODE_BIN" -e 'process.exit(Number(process.versions.node.split(".")[0]) >= 22 ? 0 : 1)'; then
  print 'Margin needs Node.js 22 or later. Update Node.js, then open this launcher again.'
  margin_pause
  exit 1
fi
if [[ ! -d node_modules/diff ]]; then
  print 'Install Margin’s dependencies first. In Terminal, run:'
  print -r -- "  cd ${(q)PWD}"
  print '  npm ci'
  print 'Then follow docs/word-add-in.md and open this launcher again.'
  margin_pause
  exit 1
fi

print 'Margin · Word and local web app'
print 'Keep this window open while reviewing. Press Control-C to stop.'
print 'If another Margin server is running, stop it before starting this one.'
print ''

"$MARGIN_NODE_BIN" scripts/word-setup.mjs start
MARGIN_LAUNCH_EXIT=$?
if (( MARGIN_LAUNCH_EXIT != 0 )); then
  print ''
  print 'Margin did not start. The message above explains what needs attention.'
  print 'For first-time Word setup, follow docs/word-add-in.md.'
  print 'To check your setup in Terminal:'
  print -r -- "  cd ${(q)PWD}"
  print '  npm run word:setup'
  margin_pause
fi
exit "$MARGIN_LAUNCH_EXIT"
