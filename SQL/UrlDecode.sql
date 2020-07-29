USE [Flow]
GO

/****** Object:  UserDefinedFunction [dbo].[UrlDecode]    Script Date: 2020/7/29 下午 04:48:39 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[UrlDecode](@url varchar(3072))
RETURNS varchar(3072)
AS
BEGIN 
    DECLARE @count int, @c char(1), @cenc char(2), @i int, @urlReturn varchar(3072) 
    SET @count = Len(@url) 
    SET @i = 1 
    SET @urlReturn = '' 
    WHILE (@i <= @count) 
     BEGIN 
        SET @c = substring(@url, @i, 1) 
        IF @c LIKE '[!%]' ESCAPE '!' 
         BEGIN 
            SET @cenc = substring(@url, @i + 1, 2) 
            SET @c = CHAR(CASE WHEN SUBSTRING(@cenc, 1, 1) LIKE '[0-9]' 
                                THEN CAST(SUBSTRING(@cenc, 1, 1) as int) 
                                ELSE CAST(ASCII(UPPER(SUBSTRING(@cenc, 1, 1)))-55 as int) 
                            END * 16 + 
                            CASE WHEN SUBSTRING(@cenc, 2, 1) LIKE '[0-9]' 
                                THEN CAST(SUBSTRING(@cenc, 2, 1) as int) 
                                ELSE CAST(ASCII(UPPER(SUBSTRING(@cenc, 2, 1)))-55 as int) 
                            END) 
            SET @urlReturn = @urlReturn + @c 
            SET @i = @i + 2 
         END 
        ELSE 
         BEGIN 
            SET @urlReturn = @urlReturn + @c 
         END 
        SET @i = @i +1 
     END 
    RETURN @urlReturn
END
GO

