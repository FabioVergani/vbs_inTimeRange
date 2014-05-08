Function inTimeRange(x,y,z)
 Dim a,b,c,r
 a=TimeValue(x)
 b=TimeValue(CDate(z))
 c=TimeValue(CDate(y))
 r=False
 if (a > c And b > a) Or (a > c And b > a)  Then
	r=True
 End if
 inTimeRange=r
End Function

Response.Write inTimeRange(Now,"15:00:00","15:19:00")
Response.Write inTimeRange(Now,"15:00:00","15:20:50")
