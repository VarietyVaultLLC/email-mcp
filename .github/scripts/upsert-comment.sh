#!/usr/bin/env bash
# upsert-comment.sh PR PREFIX BODY_FILE: post BODY_FILE as the one comment on PR
# that starts with PREFIX, editing it in place once it exists. Not
# `gh pr comment --edit-last`: review and security both comment as
# github-actions[bot], so "last" is whichever one finished second.
set -euo pipefail
pr=$1 prefix=$2 body=$3
export GH_REPO=${GH_REPO:-$GITHUB_REPOSITORY}
ids=$(gh api --paginate "repos/$GH_REPO/issues/$pr/comments" \
  --jq ".[] | select(.user.login == \"github-actions[bot]\" and (.body | startswith(\"$prefix\"))) | .id")
id=${ids%%$'\n'*}
if [ -n "$id" ]; then
  gh api -X PATCH "repos/$GH_REPO/issues/comments/$id" -F "body=@$body" > /dev/null
else
  gh pr comment "$pr" --body-file "$body"
fi
