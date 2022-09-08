Use [master]
Go
Create Database dbProducts
Go
Use dbProducts
Go
Create Table Products
(
Id int Identity(1,1) Primary Key,
ProductCode nvarchar(10),
ProductName nvarchar(20),
Barcode nvarchar(20)
)
Go
Insert Into Products (ProductCode,ProductName,Barcode) Values ('001','Fried Rice','838499304023')
Insert Into Products (ProductCode,ProductName,Barcode) Values ('002','Chicken BBQ','264738283839')
Insert Into Products (ProductCode,ProductName,Barcode) Values ('003','Egg Buritto','164338203839')
Insert Into Products (ProductCode,ProductName,Barcode) Values ('004','Maltina Beer','737637283839')
Go
Select * From Products