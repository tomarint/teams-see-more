find . -name '.DS_Store' -type f -ls -delete
rm upload.zip
cd dist
zip -r ../upload.zip ./
cd ..
unzip -l upload.zip
