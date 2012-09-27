create table #a (a int ,b varchar(max),c datetime)
insert into #a values (1,'this is a cell','2009-12-08T18:19:17.66')
insert into #a values (1,'this is another cell','2009-12-08T18:19:17.66')
insert into #a values (1,'this is another cell','2009-12-08T18:19:17.66')
insert into #a values (1,'this is another cell','2009-12-08T18:19:17.66')
insert into #a values (1,'this is another cell','2009-12-08T18:19:17.66')
create table #b (a varchar(max) ,b varchar(max))
insert into #b values ('1','this is a new sheet')
insert into #b values ('1a','this is also on hte new sheet')
insert into #b values ('2009-01-01T00:00:00','this is also on hte new sheet')
insert into #b values ('12/31/2009','this is also on hte new sheet')
create table #c (a varchar(max) ,b varchar(max))

select * from #a
select * from #c
select * from #b