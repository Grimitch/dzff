CREATE TABLE Sales (
    Id INT PRIMARY KEY,
    ProductName NVARCHAR(100),
    Category NVARCHAR(50),
    Price DECIMAL(10,2),
    Quantity INT,
    SaleDate DATETIME
);
