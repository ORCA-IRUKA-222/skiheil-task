#!/bin/sh
# GAS を実機にデプロイせずに、サーバー側のロジックだけ Node で検証する。
#   sh apps/festival-pos/test/run.sh
set -e
dir=$(dirname "$0")
node "$dir/pricing.test.js"
node "$dir/idempotency.test.js"
echo "すべてのテストが通りました"
