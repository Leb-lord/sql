CREATE ASSEMBLY VarianceAssembly
FROM 'C:\aggregate_Vaiance\aggregate_Vaiance\bin\Debug\aggregate_Vaiance.dll'
WITH PERMISSION_SET = SAFE;

CREATE AGGREGATE dbo.CalcVariance
(
    @value FLOAT
)
RETURNS FLOAT
EXTERNAL NAME VarianceAssembly.VarianceAggregate;

SELECT dbo.CalcVariance(realNumbers) AS VarianceValue
FROM dbo.aggregate_test;

insert into aggregate_test values (1.03),(53),(9.4),(32);
EXEC sp_configure 'clr enabled', '1'
RECONFIGURE;

EXEC sp_configure 'clr strict security', '0'
RECONFIGURE;

///////////////////////////////////////////////////
create assembly addToOdd
from 'C:\AddOne\AddOne\bin\Debug\AddOne.dll'
WITH PERMISSION_SET = SAFE;

CREATE FUNCTION dbo.addOne
(
    @value INT
)
RETURNS INT
AS EXTERNAL NAME addToOdd.[addOneToOdd].[addOne];

select dbo.addOne(43524523) as Result;

//////////////////////////////////////////////////
USE functionsTask;
EXEC sys.sp_cdc_enable_db;

EXEC sys.sp_cdc_enable_table 
    @source_schema = N'dbo', 
    @source_name   = N'aggregate_test', 
    @role_name     = NULL;

SELECT * FROM cdc.dbo_aggregate_test_CT;


EXEC sys.sp_cdc_cleanup_change_table 
@capture_instance = 'dbo_aggregate_test', 
@low_water_mark   = 'MIN'

insert into dbo.aggregate_test values (34); 

