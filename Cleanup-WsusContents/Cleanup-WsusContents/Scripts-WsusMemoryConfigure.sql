exec sp_configure 'min server memory' 
go

exec sp_configure 'show advanced options',1
go
reconfigure
go
exec sp_configure 'min server memory',2048
go
reconfigure
go

exec sp_configure 'min server memory' 
go