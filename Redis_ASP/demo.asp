<%
'Project: ASP Redis Client
'Author: jiezhe.net

Dim Redis
Set Redis = Server.CreateObject("vRedis.Client")

'连接Redis服务器，"Redis服务器IP","端口"
Redis.Connect "127.0.0.1","6379"

'登录密码
'Redis.Auth "foobared"

'指定的库索引号
'Redis.Select 0

'设置一个键值 "键名","值","设置键名的过期时间，单位时毫秒，不设置则填Empty",布尔值，填false时键名不存在才会设置值，不设置则填Empty
Redis.SetKeyValue "name","key",Empty,Empty

'获取键名对应的值
Response.write Redis.GetValue("name")

'如果键名已经存在，并且原来的键值为字符串，那么这个命令会把新键值追加到原来键值的结尾。如果键名不存在，那么它将首先创建一个空字符串，再执行追加操作，这种情况 APPEND 将类似于 SET 操作
'Redis.Append "name","key123"

'返回键名是否存在
'Response.write Redis.CheckExistentKey("name")

'删除键
'Redis.DeleteKey "name"

'Redis.DeleteMultipleKeys "name","key"

'设置键名的过期时间，超过时间后，将会自动删除该键名。单位秒
'Redis.Expire "name",10

'设置键名的过期时间，超过时间后，将会自动删除该键名。EXPIREAT 的作用和 EXPIRE类似，都用于为键名设置生存时间。不同在于 EXPIREAT 命令接受的时间参数是 UNIX 时间戳 Unix timestamp
'Redis.ExpireAt "name",Unix_timestamp(now()+1,false)

'单位毫秒
'Redis.PExpire "name",10
'Redis.PExpireAt "name",Unix_timestamp(now()+1,true)

'更多详细内容请参考Redis官网的文档 https://redis.io/
'Response.write Redis.Dump("name")
'Response.write Redis.Echo("china")
'Response.write Redis.GetTime()
'Response.write Redis.Ping()

'关闭当前与redis服务的连接
Redis.Quit

'获取时间戳(时间,是否精确到毫秒级)
Function Unix_timestamp(ts,IsMilli)
	If ts="" Then ts=now()
	Unix_timestamp=DateDiff("s","1970-01-01 08:00:00",ts)
	If IsMilli Then Unix_timestamp=Unix_timestamp*1000+Int(CDbl(Timer())*1000)
End Function
%>