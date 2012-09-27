create table #a (a int identity(1,1), a1 int, b uniqueidentifier, c datetime, d uniqueidentifier
	, e uniqueidentifier, f uniqueidentifier, g uniqueidentifier, h uniqueidentifier
	, i uniqueidentifier, j uniqueidentifier, k uniqueidentifier
)
insert into #a values (
		1
		, 'F9CCFB05-B5DE-4155-B207-99A60715EECE'
		, getdate()
		, 'E92D6B66-1BA5-49C3-B486-BCD5D1418E78'
		, 'CDEE3146-0665-4C53-8752-3509A615AC33'
		, '142A4FEF-19A0-484F-93FC-70D7A16FABF7'
		, 'BB29BD7B-1907-4C1A-BBCD-E721C503C084'
		, 'DF58E6E9-C694-4BD6-A8CE-2FF4778A9A38'
		, '8BCE720A-FDF3-46BA-9C58-1261CC1D5DEB'
		, 'C6EFE501-3D93-4515-A754-6AF87730FB4D'
		, 'CB4ABB09-8043-42CF-873C-5A17708AD634'
	)

set nocount on
declare @a int
set @a = 2
-- 18: 10mb, 19: 21mb, 21: 87mb, 22: 350mb, 26: 3000mb
while(@a < 10)
begin
	insert into #a
	select @a,b,c,d,e,f,g,h,i,j,k
	from #a
	set @a = @a + 1
end

-- exceeds every 105 rows
select * from #a
order by 1

-- exceeds every 105 rows
select * from #a
order by 1

-- doesn't exceed
select top 90 * from #a
order by 1

-- doesn't exceed
select top 103 * from #a
order by 1

-- exceeds every 105 rows
select * from #a
order by 1

/*

create table #a (a int identity(1,1), a1 int, b uniqueidentifier, c datetime, d uniqueidentifier
	, e uniqueidentifier, f uniqueidentifier, g uniqueidentifier, h uniqueidentifier
	, i uniqueidentifier, j uniqueidentifier, k uniqueidentifier
)
insert into #a values (
		1
		, 'F9CCFB05-B5DE-4155-B207-99A60715EECE'
		, getdate()
		, 'E92D6B66-1BA5-49C3-B486-BCD5D1418E78'
		, 'CDEE3146-0665-4C53-8752-3509A615AC33'
		, '142A4FEF-19A0-484F-93FC-70D7A16FABF7'
		, 'BB29BD7B-1907-4C1A-BBCD-E721C503C084'
		, 'DF58E6E9-C694-4BD6-A8CE-2FF4778A9A38'
		, '8BCE720A-FDF3-46BA-9C58-1261CC1D5DEB'
		, 'C6EFE501-3D93-4515-A754-6AF87730FB4D'
		, 'CB4ABB09-8043-42CF-873C-5A17708AD634'
	)
	
select OBJECT_ID, OBJECT_NAME(object_id),* from tempdb.sys.indexes
where type = 0

set nocount on
declare @a int
select @a = max(a1) + 1 from #a 

	insert into #a
	select @a,b,c,d,e,f,g,h,i,j,k
	from #a
	select @a

select 
	object_name = object_name(i.object_id),i.index_id, index_name = i.name
	, i.is_disabled
	, sum(a.total_pages) as totalPages
	, sum(a.used_pages) as usedPages
	, sum(a.data_pages) as dataPages
	, (sum(a.total_pages) * 8) / 1024 as totalSpaceMB
	, (sum(a.used_pages) * 8) / 1024 as usedSpaceMB
	, (sum(a.data_pages) * 8) / 1024 as dataSpaceMB
from 
	sys.indexes i with(NOLOCK)
	left join sys.partitions p with(NOLOCK) on i.object_id = p.object_id
		and i.index_id = p.index_id
	left join sys.allocation_units a with(NOLOCK) 
		on p.partition_id = a.container_id
where 
	i.object_id = 873025441
group by
	object_name(i.object_id),i.index_id, i.name
	, i.fill_factor, i.is_disabled
order by 
	(sum(a.used_pages) * 8) / 1024 desc

-- for 17
select rows=count(*),max_ident=max(a) from #a
--rows	max_ident
--32768	32768

-- for 18
select rows=count(*),max_ident=max(a) from #a
--rows	max_ident
--65536	65536

-- for 19
select rows=count(*),max_ident=max(a) from #a
--rows	max_ident
--131072	131072

-- for 25
select rows=count(*),max_ident=max(a) from #a
--rows	max_ident
--16777216	16777216

*/