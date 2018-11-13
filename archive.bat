del latest.zip
7z\7za a -x@7z\.ignore latest.zip *.* * -r
call aws lambda update-function-code --function-name excel_py --zip-file fileb://latest.zip --profile cs_deployLambda
pause
