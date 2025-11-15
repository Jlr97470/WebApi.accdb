Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7937
    DatasheetFontHeight =11
    ItemSuffix =8
    Left =-32086
    Top =1530
    Right =-19831
    Bottom =16695
    RecSrcDt = Begin
        0xb9f2a7a2ce14e640
    End
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =4535
            Name ="Détail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =5101
                    Top =793
                    Width =1686
                    Height =315
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="CmbFromCurrencyName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Currency].[CurrencyName], Currency.[CurrencyLabel] FROM [Currency] ORDER"
                        " BY [CurrencyName]; "
                    ColumnWidths ="1441;1441"
                    GridlineColor =10921638

                    LayoutCachedLeft =5101
                    LayoutCachedTop =793
                    LayoutCachedWidth =6787
                    LayoutCachedHeight =1108
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2830
                            Top =793
                            Width =2295
                            Height =320
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="CmbFromCurrencyName_Étiquette"
                            Caption ="From Currency Name"
                            GridlineColor =10921638
                            LayoutCachedLeft =2830
                            LayoutCachedTop =793
                            LayoutCachedWidth =5125
                            LayoutCachedHeight =1113
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =5107
                    Top =2267
                    Width =1686
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="CmbToCurrencyName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Currency].[CurrencyName], Currency.[CurrencyLabel] FROM [Currency] ORDER"
                        " BY [CurrencyName]; "
                    ColumnWidths ="1442;1442"
                    GridlineColor =10921638

                    LayoutCachedLeft =5107
                    LayoutCachedTop =2267
                    LayoutCachedWidth =6793
                    LayoutCachedHeight =2582
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2821
                            Top =2267
                            Width =2310
                            Height =320
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="CmbToCurrencyName_Étiquette"
                            Caption ="To Currency Name"
                            GridlineColor =10921638
                            LayoutCachedLeft =2821
                            LayoutCachedTop =2267
                            LayoutCachedWidth =5131
                            LayoutCachedHeight =2587
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1712
                    Top =793
                    Width =1131
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TxtCurrency"
                    Format ="General Number"
                    GridlineColor =10921638

                    LayoutCachedLeft =1712
                    LayoutCachedTop =793
                    LayoutCachedWidth =2843
                    LayoutCachedHeight =1108
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =793
                            Width =1150
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="TxtCurrency_Étiquette"
                            Caption ="Currency"
                            GridlineColor =10921638
                            LayoutCachedLeft =566
                            LayoutCachedTop =793
                            LayoutCachedWidth =1716
                            LayoutCachedHeight =1108
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    FilterLookup =0
                    IMESentenceMode =3
                    Left =1712
                    Top =2267
                    Width =1131
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TxtCurrencyC"
                    Format ="General Number"
                    GridlineColor =10921638

                    LayoutCachedLeft =1712
                    LayoutCachedTop =2267
                    LayoutCachedWidth =2843
                    LayoutCachedHeight =2582
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =2267
                            Width =1150
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="TxtCurrencyC_Étiquette"
                            Caption ="Currency C."
                            GridlineColor =10921638
                            LayoutCachedLeft =566
                            LayoutCachedTop =2267
                            LayoutCachedWidth =1716
                            LayoutCachedHeight =2582
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =566
                    Top =3401
                    Width =2580
                    Height =420
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnConvert"
                    Caption ="Convert"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =3401
                    LayoutCachedWidth =3146
                    LayoutCachedHeight =3821
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                End
                Begin Label
                    OverlapFlags =85
                    Left =3551
                    Top =3461
                    Width =3810
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblStatus"
                    Caption ="Done"
                    GridlineColor =10921638
                    LayoutCachedLeft =3551
                    LayoutCachedTop =3461
                    LayoutCachedWidth =7361
                    LayoutCachedHeight =3776
                End
            End
        End
    End
End
CodeBehindForm
' See "Currency.cls"
