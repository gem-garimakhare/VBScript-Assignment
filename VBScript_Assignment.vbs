function cubeOfNumber(num)
	cubeOfNumber = num*num*num
	'msgbox(cubeOfNumber)
end function

function isArmstrong(n)
	dim temp,sum,r
	temp=n
	sum=0
	while temp>0
	r=temp mod 10
	sum=sum+cubeOfNumber(r)
	temp=temp\10
	wend
	if sum=n then
	isArmstrong=true
	else
	isArmstrong=false
	end if
end function

function generateArmstrongNumbers(range_s,range_e)
	for i=cint(range_s) to range_e
		'msgbox(i)
		res = isArmstrong(i)
		if res=true then
			msgbox(cstr(i)+" is an armstrong number")
		end if
	next
end function

dim range_s,range_e,res
range_s = inputbox("Enter value for range start")
range_e = inputbox("Enter value for range end")

StartTime = Timer()
generateArmstrongNumbers range_s,range_e
EndTime = Timer()
timeTaken = EndTime-StartTime
msgbox("Time taken to generate armstrong numbers between given range is: " & FormatNumber(timeTaken, 2))





