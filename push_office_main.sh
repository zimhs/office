#!/bin/zsh
cd "$(dirname "$0")"
git push https://github.com/zimhs/office.git HEAD:main
echo "exit=$?"
