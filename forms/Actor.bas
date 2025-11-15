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
    Width =9637
    DatasheetFontHeight =11
    ItemSuffix =32
    Left =-31726
    Top =1335
    Right =-8191
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
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
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
            Height =10771
            Name ="Détail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =566
                    Top =5839
                    Width =2580
                    Height =420
                    TabIndex =1
                    ForeColor =4210752
                    Name ="CmdListe"
                    Caption ="Liste"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =5839
                    LayoutCachedWidth =3146
                    LayoutCachedHeight =6259
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =566
                    Top =1133
                    Width =8511
                    Height =4530
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="CmbActor"
                    RowSourceType ="Value List"
                    ColumnWidths ="567;1701;1701;1701"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =566
                    LayoutCachedTop =1133
                    LayoutCachedWidth =9077
                    LayoutCachedHeight =5663
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =566
                            Top =566
                            Width =1150
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="CmbActor_Étiquette"
                            Caption ="Actor"
                            GridlineColor =10921638
                            LayoutCachedLeft =566
                            LayoutCachedTop =566
                            LayoutCachedWidth =1716
                            LayoutCachedHeight =881
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =566
                    Top =6803
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TxtActor_id"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =6803
                    LayoutCachedWidth =2267
                    LayoutCachedHeight =7118
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =566
                            Top =6519
                            Width =1695
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="TxtActor_id_Étiquette"
                            Caption ="Actor_id"
                            GridlineColor =10921638
                            LayoutCachedLeft =566
                            LayoutCachedTop =6519
                            LayoutCachedWidth =2261
                            LayoutCachedHeight =6834
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =2267
                    Top =6803
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TxtFirst_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =2267
                    LayoutCachedTop =6803
                    LayoutCachedWidth =3968
                    LayoutCachedHeight =7118
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =2267
                            Top =6519
                            Width =1695
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="TxtFirst_Name_Étiquette"
                            Caption ="First_name"
                            GridlineColor =10921638
                            LayoutCachedLeft =2267
                            LayoutCachedTop =6519
                            LayoutCachedWidth =3962
                            LayoutCachedHeight =6834
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =3968
                    Top =6803
                    Height =315
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TxtLast_name"
                    GridlineColor =10921638

                    LayoutCachedLeft =3968
                    LayoutCachedTop =6803
                    LayoutCachedWidth =5669
                    LayoutCachedHeight =7118
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =3968
                            Top =6519
                            Width =1695
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="TxtLast_name_Étiquette"
                            Caption ="Last_Name"
                            GridlineColor =10921638
                            LayoutCachedLeft =3968
                            LayoutCachedTop =6519
                            LayoutCachedWidth =5663
                            LayoutCachedHeight =6834
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =566
                    Top =7370
                    Width =2580
                    Height =420
                    TabIndex =6
                    ForeColor =4210752
                    Name ="CmdUpdate"
                    Caption ="Mise a jour"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =7370
                    LayoutCachedWidth =3146
                    LayoutCachedHeight =7790
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5669
                    Top =6803
                    Width =3411
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TxtLast_Update"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=Now()"
                    InputMask ="00/00/0000\\ 00:00:00;0;_"
                    GridlineColor =10921638

                    LayoutCachedLeft =5669
                    LayoutCachedTop =6803
                    LayoutCachedWidth =9080
                    LayoutCachedHeight =7118
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =5669
                            Top =6519
                            Width =3405
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="TxtLast_Update_Étiquette"
                            Caption ="Last_Update"
                            GridlineColor =10921638
                            LayoutCachedLeft =5669
                            LayoutCachedTop =6519
                            LayoutCachedWidth =9074
                            LayoutCachedHeight =6834
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =2267
                    Top =8221
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TxtFirst_NameAdd"
                    GridlineColor =10921638

                    LayoutCachedLeft =2267
                    LayoutCachedTop =8221
                    LayoutCachedWidth =3968
                    LayoutCachedHeight =8536
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2267
                            Top =7937
                            Width =1695
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="TxtFirst_NameAdd_Étiquette"
                            Caption ="First_name"
                            GridlineColor =10921638
                            LayoutCachedLeft =2267
                            LayoutCachedTop =7937
                            LayoutCachedWidth =3962
                            LayoutCachedHeight =8252
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =3968
                    Top =8221
                    Height =315
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TxtLast_nameAdd"
                    GridlineColor =10921638

                    LayoutCachedLeft =3968
                    LayoutCachedTop =8221
                    LayoutCachedWidth =5669
                    LayoutCachedHeight =8536
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =3968
                            Top =7937
                            Width =1695
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="TxtLast_nameAdd_Étiquette"
                            Caption ="Last_Name"
                            GridlineColor =10921638
                            LayoutCachedLeft =3968
                            LayoutCachedTop =7937
                            LayoutCachedWidth =5663
                            LayoutCachedHeight =8252
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =566
                    Top =8788
                    Width =2580
                    Height =420
                    TabIndex =10
                    ForeColor =4210752
                    Name ="CmdAdd"
                    Caption ="Ajout"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =8788
                    LayoutCachedWidth =3146
                    LayoutCachedHeight =9208
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5669
                    Top =8221
                    Width =3411
                    Height =315
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TxtLast_UpdateAdd"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=now()"
                    InputMask ="00/00/0000\\ 00:00:00;0;_"
                    GridlineColor =10921638

                    LayoutCachedLeft =5669
                    LayoutCachedTop =8221
                    LayoutCachedWidth =9080
                    LayoutCachedHeight =8536
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =5669
                            Top =7937
                            Width =3405
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="TxtLast_UpdateAdd_Étiquette"
                            Caption ="Last_Update"
                            GridlineColor =10921638
                            LayoutCachedLeft =5669
                            LayoutCachedTop =7937
                            LayoutCachedWidth =9074
                            LayoutCachedHeight =8252
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =566
                    Top =9637
                    Width =2580
                    Height =420
                    TabIndex =11
                    ForeColor =4210752
                    Name ="CmdDelete"
                    Caption ="Supprimer"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =9637
                    LayoutCachedWidth =3146
                    LayoutCachedHeight =10057
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                End
                Begin TextBox
                    Locked = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3571
                    Top =5725
                    Width =5520
                    Height =690
                    TabIndex =12
                    ForeColor =6710886
                    Name ="lblStatus"
                    GridlineColor =10921638

                    LayoutCachedLeft =3571
                    LayoutCachedTop =5725
                    LayoutCachedWidth =9091
                    LayoutCachedHeight =6415
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
                Begin TextBox
                    Locked = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3571
                    Top =7200
                    Width =5520
                    Height =630
                    TabIndex =13
                    ForeColor =6710886
                    Name ="lblStatusUpdate"
                    GridlineColor =10921638

                    LayoutCachedLeft =3571
                    LayoutCachedTop =7200
                    LayoutCachedWidth =9091
                    LayoutCachedHeight =7830
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
                Begin TextBox
                    Locked = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3571
                    Top =8730
                    Width =5520
                    Height =615
                    TabIndex =14
                    ForeColor =6710886
                    Name ="lblStatusAdd"
                    GridlineColor =10921638

                    LayoutCachedLeft =3571
                    LayoutCachedTop =8730
                    LayoutCachedWidth =9091
                    LayoutCachedHeight =9345
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
                Begin TextBox
                    Locked = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3571
                    Top =9524
                    Width =5520
                    Height =615
                    TabIndex =15
                    ForeColor =6710886
                    Name ="lblStatusDelete"
                    GridlineColor =10921638

                    LayoutCachedLeft =3571
                    LayoutCachedTop =9524
                    LayoutCachedWidth =9091
                    LayoutCachedHeight =10139
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
            End
        End
    End
End
CodeBehindForm
' See "Actor.cls"
