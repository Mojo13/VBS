dim count
dim Text
count = 0
Text = "This is the count"

while count < 10
	wscript.echo Text & count
	count = count + 1
rem 	wscript.sleep 100
wend