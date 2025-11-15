CREATE TABLE [Language] (
  [LanguageNum] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [LanguageCode] VARCHAR (255),
  [LanguageLabel] VARCHAR (255)
)
