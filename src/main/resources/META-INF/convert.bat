rem see https://developer.github.com/v3/markdown/

wget --header="Content-Type: text/plain" --debug --verbose --no-check-certificate --output-document=description.html  --post-file description.md https://api.github.com/markdown/raw 