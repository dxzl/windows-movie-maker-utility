�
 TFORMMAIN 0~  TPF0	TFormMainFormMainLeft Top� ActiveControlMemo1Caption+Windows Live Movie Maker Utility 2.0 by DTSClientHeightVClientWidthVColor	clBtnFaceConstraints.MinHeight|Constraints.MinWidthbFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	Icon.Data
6           (  &          �  N  (                �                         �  �   �� �   � � ��  ��� ���   �  �   �� �   � � ��  ��� ����˻UU����˻UU����˻UU����˻UU����˻UU����˻UU����˻UU����˻UU����˻UU����˻UU����˻UU����˻UU����˻UU����˻UU����˻UU����˻UU                                                                (       @         �                        �  �   �� �   � � ��  ��� ���   �  �   �� �   � � ��  ��� ����� ��     DD ����� 	��     @����� 
���  @  �U��� 
���� @  �U����
���� @  �UU���
�����@  ��U\�� �����  @��UU�� ����� DD �UUU� �����    �UUU� �����    �UUU\�
����    �UUU\�
����     �UUU\�
����     �UUUU�
���      ��UUU� ���      ��UUU� ��       �UUU� ��       �UUU� 	�       �UUU� 	�       �UU\� �       �UU\� ��       �UU\� ��       �UU�� ���      ��U��  ���      ��\�"  ����     ���"" �����    ���"" ������   ��""" �����    ��                                                             ��  ?�  F  B  Z  B   �   À  ��  ��  �� �� �� �� �� �� �� �� �� �� �� �� ��  ��  �  �  �  �  �  ��������Menu	MainMenu1PositionpoDesktopCenterOnCreate
FormCreate	OnDestroyFormDestroyOnShowFormShow
TextHeight TPanelPanel1Left Top WidthVHeight9AlignalTop
BevelOuter	bvLoweredTabOrderExplicitWidthR TButtonButtonReadProjectFileLeftTopWidth� HeightHint,Read a windows movie-maker wlmp project-fileBiDiModebdLeftToRightCaption1. Read .wlmp project-fileFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style ParentBiDiMode
ParentFontParentShowHintShowHint	TabOrder OnClickButtonReadProjectFileClick  TButtonButtonSelectMediaFolderLeftTopWidth�HeightHint,Read a windows movie-maker wlmp project-fileCaptionI2. Select folder where your image and movie/image/audio files are locatedFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontParentShowHintShowHint	TabOrderOnClickButtonSelectMediaFolderClick  TButton
ButtonHelpLeft�TopWidthbHeightHintSay no more ;) ;) *nod* *nod*CaptionHelpFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontParentShowHintShowHintTabOrderOnClickButtonHelpClick  TButtonButtonApplyNewRootPathLeft� TopWidthHeightHintChange the common root pathCaption3. Apply new root-project pathFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontParentShowHintShowHint	TabOrderOnClickButtonApplyNewRootPathClick  TButtonButtonSaveFileLeft�TopWidthbHeightHint'Save the edit window contents as a fileCaption4. Save fileFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontParentShowHintShowHint	TabOrderOnClickButtonSaveFileClick   TMemoMemo1Left Top9WidthVHeight� HintEdit windowAlignalClientFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont
ScrollBarsssBothTabOrder WordWrapExplicitWidthRExplicitHeight�   	TGroupBox	GroupBox1Left TopWidthVHeight@AlignalBottomTabOrderExplicitTopExplicitWidthR TLabelLabelProjectFilePathLeftTop.WidthRHeightAlignalBottomFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontExplicitWidth  TLabelLabelMediaFolderPathLeftTopWidthRHeightAlignalBottomFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontParentShowHintShowHint	LayouttlBottomExplicitWidth  TProgressBarProgressBar1LeftTopWidthRHeight
AlignalTopStepTabOrder ExplicitWidthN   TOpenDialogOpenDialog1
DefaultExtwlmp
InitialDir%HOMEPATH%\Documents\MovieMakerLeftPTopX  TSaveDialogSaveDialog1
DefaultExtwlmp
InitialDir%HOMEPATH%\Documents\MovieMakerLeft� TopX  	TMainMenu	MainMenu1Left�TopX 	TMenuItem	MenuToolsCaptionTools 	TMenuItemCopyMovieImageFilesToNewFolder1Caption8Copy a working project and its media files to new folderHintHCopy the movies/images used in your project to a folder that you specifyShortCutC�  OnClick$CopyMovieImageFilesToNewFolder1Click  	TMenuItemMoveprojecttonewfolder1Caption8Move a working project and its media files to new folderHintHMove the movies/images used in your project to a folder that you specifyShortCutM�  OnClickMoveprojecttonewfolder1Click   	TMenuItem	MenuResetCaptionResetHintClear list of search foldersOnClickMenuResetClick   TTimerTimerSearchTimeoutEnabledInterval'OnTimerTimerSearchTimeoutTimerLeft� TopX   