--Insert into [DESKTOP-JBRFH9E].[testdb].[dbo].[F2ASC] select * from [49.50.103.132].[ASTROLOGYSOFTWARE_DB].[dbo].[F2ASC]
--Insert into [DESKTOP-JBRFH9E].[testdb].[dbo].[F2BASE] select * from [49.50.103.132].[ASTROLOGYSOFTWARE_DB].[dbo].[F2BASE]
--Insert into [DESKTOP-JBRFH9E].[testdb].[dbo].[F2PLANETS] select 'ME',p1,p2,p3,p4,p5,p6,p7 from [DESKTOP-JBRFH9E].[testdb].[dbo].[F2ASC]
--select count(*) from [DESKTOP-JBRFH9E].[testdb].[dbo].[F2PLANETS]
--Select * into [HCUSP] from [49.50.103.132].[ASTROLOGYSOFTWARE_DB].[dbo].[TABLE_HCUSP]
--Select * into [HPLANET] from [49.50.103.132].[ASTROLOGYSOFTWARE_DB].[dbo].[TABLE_HPLANET]
--SELECT * INTO CUSP FROM [49.50.103.132].[ASTROLOGYSOFTWARE_DB].[dbo].[CUSP]
--TRUNCATE table cusp
--SELECT * INTO HRAKE FROM [49.50.103.132].[ASTROLOGYSOFTWARE_DB].[dbo].[HRAKE]
--TRUNCATE table HRAKE
--insert into F2PLANETS select p0 + p1 + p2 + p3 + p4 + p5 + p6 + p7 from testdb.dbo.F2PLANETS
--update F2BASE_ set p3 = 'V'
--			 where p3 = 'VE'
--update F2BASE_ set p3 = 'R'
--			 where p3 = 'RA'
--update F2BASE_ set p3 = 'K'
--			 where p3 = 'KE'
--update F2BASE_ set p3 = 'm'
--			 where p3 = 'ME'
--update F2BASE_ set p3 = 'O'
--			 where p3 = 'MO'
--update F2BASE_ set p3 = 'S'
--			 where p3 = 'SU'
--update F2BASE_ set p3 = 's'
--			 where p3 = 'SA'
--update F2BASE_ set p3 = 'J'
--			 where p3 = 'JU'
--update F2BASE_ set p3 = 'M'
--			 where p3 = 'MA'
----select * from F2BASE_

alter table F2BASE_ alter column p0 char(1);
alter table F2BASE_ alter column p1 char(1);
alter table F2BASE_ alter column p2 char(1);
alter table F2BASE_ alter column p3 char(1);
alter table F2BASE_ alter column p4 char(1);
alter table F2BASE_ alter column p5 char(1);
alter table F2BASE_ alter column p6 char(1);
alter table F2BASE_ alter column p7 char(1);


--if exists (select * from dbo.sysobjects where name = N'sp_permutation' and xtype='P')
--drop procedure sp_permutation
--GO
--CREATE procedure sp_permutation(@str varchar(10))
--As
--declare @t varchar(10) = @str; with s(t,n)
--as 
--(select substring(@t,1,1),1 union all select substring(@t,n+1,1),n+1 from s where n<len(@t))
--,j(t) as (select cast(t as varchar(10)) from s union all select cast(j.t+s.t as varchar(10))from j,s where patindex('%'+s.t+'%',j.t)=0)
--(select t from j where len(t)=len(@t))

alter table cusp_ alter column S1 varchar(34);
