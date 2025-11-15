CREATE TABLE [Currency] (
  [CurrencyNum] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [CurrencyName] VARCHAR (3) CONSTRAINT [IndCurrencyName] UNIQUE,
  [CurrencyLabel] VARCHAR (255)
)
