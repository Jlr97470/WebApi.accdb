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
    Left =-12031
    Top =4455
    Right =-4096
    Bottom =11250
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
            Height =6803
            Name ="Détail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =566
                    Top =1133
                    Width =6246
                    Height =1695
                    Name ="TxtQuestion"

                    LayoutCachedLeft =566
                    LayoutCachedTop =1133
                    LayoutCachedWidth =6812
                    LayoutCachedHeight =2828
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =566
                            Top =566
                            Width =1150
                            Height =315
                            Name ="TxtQuestion_Étiquette"
                            Caption ="Question"
                            LayoutCachedLeft =566
                            LayoutCachedTop =566
                            LayoutCachedWidth =1716
                            LayoutCachedHeight =881
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    FilterLookup =0
                    IMESentenceMode =3
                    Left =566
                    Top =3968
                    Width =6246
                    Height =1530
                    TabIndex =1
                    Name ="TxtReponse"

                    LayoutCachedLeft =566
                    LayoutCachedTop =3968
                    LayoutCachedWidth =6812
                    LayoutCachedHeight =5498
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =566
                            Top =3401
                            Width =1125
                            Height =315
                            Name ="TxtReponse_Étiquette"
                            Caption ="Reponse"
                            LayoutCachedLeft =566
                            LayoutCachedTop =3401
                            LayoutCachedWidth =1691
                            LayoutCachedHeight =3716
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =566
                    Top =5839
                    Width =2580
                    Height =420
                    TabIndex =2
                    Name ="btnSend"
                    Caption ="Envoyer"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =566
                    LayoutCachedTop =5839
                    LayoutCachedWidth =3146
                    LayoutCachedHeight =6259
                End
                Begin Label
                    OverlapFlags =85
                    Left =3551
                    Top =5899
                    Width =3810
                    Height =315
                    Name ="lblStatus"
                    Caption ="Done"
                    LayoutCachedLeft =3551
                    LayoutCachedTop =5899
                    LayoutCachedWidth =7361
                    LayoutCachedHeight =6214
                End
            End
        End
    End
End
CodeBehindForm
' See "ChatGPT.cls"
