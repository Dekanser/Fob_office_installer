::��騥 ᢥ����� � �।�⢥ ࠧ����뢠��� https://docs.microsoft.com/ru-ru/deployoffice/overview-of-the-office-2016-deployment-tool#exclude-onedrive-when-installing-office-365-proplus-or-other-applications
::���᮪ ID �த�⮢ Microsoft https://docs.microsoft.com/ru-ru/office365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run
::� ࠧ����뢠��� �몮��� ����⮢ https://docs.microsoft.com/ru-ru/deployoffice/overview-of-deploying-languages-in-office-365-proplus
::��ࠬ���� ���䨣��樨 xml 䠩�� https://docs.microsoft.com/ru-ru/deployoffice/configuration-options-for-the-office-2016-deployment-tool
::���᮪ ID �몮� https://docs.microsoft.com/ru-ru/deployoffice/overview-of-deploying-languages-in-office-365-proplus#languages-culture-codes-and-companion-proofing-languages
::���᮪ ID �몮� 2 https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/available-language-packs-for-windows


::�᫨ ��࠭� Office ⮣�� �� �㤥� ��⠭������ ProjectPro,VisioPro �.� ��� �� �室�� � ����� Ofice365 � ����� ���� ��⠭������ ⮫쪮 �⤥�쭮
::�᫨ ��࠭� Separately ⮣�� �� �㤥� ��⠭������ OneNote, Lync, Groove, OneDrive, Teams �.� ��� �� ����� ���� ��⠭������ ���ᮡ����, ��� ��⠭���� ���
@echo off
Title Fobaz Office Installer
color 0A
mode 65,15
cd /d %~dp0
:: �������� ��⨢��� ��४��� �� ��, � ���ன ��室���� ����᪠��� 䠩�. ����室��� �⮡� �������� 䠩� setup.exe
If not exist "%cd%\setup.exe" (echo.������� ��� 䠩� � ���� ��४��� � setup.exe&echo.&echo.Put this file in the same directory as setup.exe& echo.&pause&exit)


Set Download=*& Set Install=*
Set other= & Set ru=*& Set Eng= 
Set Office=*& Set Separately= 
Set All= & Set Word=*& Set Excel=*& Set PowerPoint=*& Set Access= & Set OneNote= & Set Outlook= & Set Publisher= & Set Lync= & Set Teams= & Set ProjectPro= & Set VisioPro= & Set Groove= 
:: ������� ��६����� � ��ᢠ������ �� �஡���

echo.������஢��� xml 䠩� ��� ����ன�� ��⠭����? [y \ n ]
echo.Generate xml file for settings installation? [y \ n ]
Set /P setQ= :
If /i "%setQ%"=="n" (echo ������ ������ ��� ^(������ ���७�� ".xml"^) xml 䠩��. & echo Input full name ^(include File extension ".xml"^) xml file. & Set /P xmlFileName= :& goto :startWOs) else (Set xmlFileName=Temp.xml )
::�᫨ ������� "n" ⮣�� ��३� ����� ��� xml 䠩�� � ��३� � ":startWOs".
echo.^<Configuration^> > Temp.xml
echo.  ^<Add OfficeClientEdition="64" Channel="Monthly"^> >> Temp.xml


mode 35,35
:do
Set �heckbox_ERROR=0

If /i "%otvet%"=="d" (@If "%Download%"==" " (Set Download=*) else (Set Download= ))
If /i "%otvet%"=="i" (@If "%Install%"==" " (Set Install=*) else (Set Install= ))

If /i "%otvet%"=="0" (@If "%other%"==" " (Set other=*) else (Set other= ))
If /i "%otvet%"=="1" (@If "%ru%"==" " (Set ru=*) else (Set ru= ))
If /i "%otvet%"=="2" (@If "%Eng%"==" " (Set Eng=*) else (Set Eng= ))
If /i "%otvet%"=="3" (@If "%Office%"==" " (Set Office=*& Set Separately= ) else (Set Office= & Set Separately=*))
If /i "%otvet%"=="4" (@If "%Separately%"==" " (Set Separately=*& Set Office= ) else (Set Separately= & Set Office=*))
If /i "%otvet%"=="5" (@If "%All%"==" " (Set All=*) else (Set All= ))
:: All ����� ��।�����, �㦭� �⮡� �� ���⠢������ ��窥, �⠢����� ��窨 �� �� 祪����� �ணࠬ�, ⮣�� ����� �㤥� ���� ����� :all � :notall
:: �� ⮣�� �� ���⠢������ ����� ALL �㤥� �뫥���� ��� ������ konflikt
If /i "%otvet%"=="6" (@If "%Word%"==" " (Set Word=*) else (Set Word= ))
If /i "%otvet%"=="7" (@If "%Excel%"==" " (Set Excel=*) else (Set Excel= ))
If /i "%otvet%"=="8" (@If "%PowerPoint%"==" " (Set PowerPoint=*) else (Set PowerPoint= ))
If /i "%otvet%"=="9" (@If "%Access%"==" " (Set Access=*) else (Set Access= ))
If /i "%otvet%"=="10" (@If "%OneNote%"==" " (Set OneNote=*) else (Set OneNote= ))
If /i "%otvet%"=="11" (@If "%Outlook%"==" " (Set Outlook=*) else (Set Outlook= ))

If /i "%otvet%"=="12" (@If "%Publisher%"==" " (Set Publisher=*) else (Set Publisher= ))

If /i "%otvet%"=="13" (@If "%Lync%"==" " (Set Lync=*) else (Set Lync= ))
If /i "%otvet%"=="14" (@If "%Teams%"==" " (Set Teams=*) else (Set Teams= ))


If /i "%otvet%"=="15" (@If "%ProjectPro%"==" " (Set ProjectPro=*) else (Set ProjectPro= ))
If /i "%otvet%"=="16" (@If "%VisioPro%"==" " (Set VisioPro=*) else (Set VisioPro= ))

If /i "%otvet%"=="17" (@If "%Groove%"==" " (Set Groove=*) else (Set Groove= ))


cls
echo.
echo  �������������������������������ͻ
echo  �d [%Download%] Download		 �
echo  �i [%Install%] Install			 �
echo  �				 �
echo  �Lang				 �
echo  �0 [%other%] Add other		 �
echo  �1 [%ru%] Ru			 �
echo  �2 [%eng%] Eng			 �
echo  �				 �
echo  �mod				 �
echo  �3 [%Office%] Office			 �
echo  �4 [%Separately%] Separately		 �
echo  �				 �
echo  �Prog				 �
echo  �5 [%All%] All			 �
echo  �	or			 �
echo  �6 [%Word%] Word			 �
echo  �7 [%Excel%] Excel			 �
echo  �8 [%PowerPoint%] PowerPoint		 �
echo  �9 [%Access%] Access			 �
echo  �10 [%OneNote%] OneNote		 �
echo  �11 [%Outlook%] Outlook		 �
echo  �12 [%Publisher%] Publisher		 �

echo  �13 [%Lync%] Lync			 �
echo  �14 [%Teams%] Teams			 �

echo  �15 [%ProjectPro%] ProjectPro		 �
echo  �16 [%VisioPro%] VisioPro		 �

echo  �17 [%Groove%] Groove			 �

echo  �				 �
echo  �Type "go" for start		 �
echo  �������������������������������ͼ
Set /P otvet= :
::timeout 2 >nul
:while
If /i "%otvet%"=="go" (goto next) else (goto do)
:next
::ping -n 2 localhost >nul
GOTO END
:END

If "%other%"=="*" (Set /P LangAdd= Input Language ID:)
::���������� �몠 ������, �᫨ �⮨� 祪���� "other"
:: ���᮪ ID �몮� https://docs.microsoft.com/ru-ru/deployoffice/overview-of-deploying-languages-in-office-365-proplus#languages-culture-codes-and-companion-proofing-languages

If "%Office%"=="*" (@If "%Separately%"=="*" (echo. & echo �롥�� Office ��� Separately & Set �heckbox_ERROR=1))
:: ��ப� ��� ��ᯮ����� �.� �᫮��� ������� �� �믮������ ��⮬� �� � ��।���� 祪�����, ⥯��� �� �롮� 祪���� Office, 祪���� Separately ᭨������, �� �롮� Separately, ᭨������ Office
If "%Office%"==" " (@If "%Separately%"==" " (echo. & echo �롥�� Office ��� Separately & Set �heckbox_ERROR=1))
If "%Download%"==" " (@If "%Install%"==" " (echo. & echo �롥�� Download ��� Install & Set �heckbox_ERROR=1))
If "%�heckbox_ERROR%"=="1" (Timeout 5 & goto do)
:: ���� �� �롮� ��� �� �롮� �ࠧ� ���� �㭪⮢ � "Office" � "Separately"

If "%Office%"=="*" (goto Office) else (goto Separately)
:: ��⠭���� ��� �ணࠬ� � ���� Office365 ��� ��⠭���� ������ �ணࠬ�� �매��쭮, ��� ᢮�� ���������

:Office
echo.    ^<Product ID="O365ProPlusRetail"^> >> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)

If "%All%"=="*" (echo.    ^</Product^> >> Temp.xml & goto start)
:: �᫨ ��࠭� "All" � ������� ⥣ "Product" � ��३� � ":start"

If "%Word%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="Word" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="Word" /^>>> Temp.xml)
If "%Excel%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="Excel" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="Excel" /^>>> Temp.xml)
If "%PowerPoint%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="PowerPoint" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="PowerPoint" /^>>> Temp.xml)
If "%Access%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="Access" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="Access" /^>>> Temp.xml)
If "%OneNote%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="OneNote" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="OneNote" /^>>> Temp.xml)
If "%Outlook%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="Outlook" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="Outlook" /^>>> Temp.xml)


::echo.      ^<ExcludeApp ID="Lync" /^>>> Temp.xml
If "%Lync%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="Lync" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="Lync" /^>>> Temp.xml)
::echo.      ^<ExcludeApp ID="Publisher" /^>>> Temp.xml
If "%Publisher%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="Publisher" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="Publisher" /^>>> Temp.xml)
::echo.      ^<ExcludeApp ID="Groove" /^>>> Temp.xml
If "%Groove%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="Groove" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="Groove" /^>>> Temp.xml)
::echo.      ^<ExcludeApp ID="OneDrive" /^>>> Temp.xml
::If "%OneDrive%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="OneDrive" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="OneDrive" /^>>> Temp.xml)
::echo.      ^<ExcludeApp ID="Teams" /^>>> Temp.xml
If "%Teams%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="Teams" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="Teams" /^>>> Temp.xml)





echo.      ^<ExcludeApp ID="Project" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="Visio" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="Skype" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="Skypeforbusiness" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="OneDriveforbusiness" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="InfoPath" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="SharePointDesigner" /^>>> Temp.xml

echo.    ^</Product^> >> Temp.xml

goto start




:Separately
echo ��� �㭪樮��� ��� �� �������!
If "%All%"=="*" (goto all) else (goto notall)

goto start




:all

echo ��� �㭪樮��� ��� �� �������!
echo.    ^<Product ID="Word2019Retail"^>>> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)
echo.    ^</Product^>>> Temp.xml

echo.    ^<Product ID="Excel2019Retail"^>>> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)
echo.    ^</Product^>>> Temp.xml

echo.    ^<Product ID="PowerPoint2019Retail"^>>> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)
echo.    ^</Product^>>> Temp.xml

echo.    ^<Product ID="Access2019Retail"^>>> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)
echo.    ^</Product^>>> Temp.xml

echo.    ^<Product ID="Outlook2019Retail"^>>> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)
echo.    ^</Product^>>> Temp.xml




echo.    ^<Product ID="Publisher2019Retail"^>>> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)
echo.    ^</Product^>>> Temp.xml

echo.    ^<Product ID="ProjectPro2019Retail"^>>> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)
echo.    ^</Product^>>> Temp.xml

echo.    ^<Product ID="VisioPro2019Retail"^>>> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)
echo.    ^</Product^>>> Temp.xml



goto start

:notall

echo ��� �㭪樮��� ��� �� �������!

If "%Word%"=="*" (echo.    ^<Product ID="Word2019Retail"^>>> Temp.xml & (@If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)) & (@If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)) & (@If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)) & echo.    ^</Product^>>> Temp.xml)

If "%Excel%"=="*" (echo.    ^<Product ID="Excel2019Retail"^>>> Temp.xml & (@If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)) & (@If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)) & (@If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)) & echo.    ^</Product^>>> Temp.xml)

If "%PowerPoint%"=="*" (echo.    ^<Product ID="PowerPoint2019Retail"^>>> Temp.xml & (@If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)) & (@If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)) & (@If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)) & echo.    ^</Product^>>> Temp.xml)

If "%Access%"=="*" (echo.    ^<Product ID="Access2019Retail"^>>> Temp.xml & (@If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)) & (@If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)) & (@If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)) & echo.    ^</Product^>>> Temp.xml)

If "%Outlook%"=="*" (echo.    ^<Product ID="Outlook2019Retail"^>>> Temp.xml & (@If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)) & (@If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)) & (@If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)) & echo.    ^</Product^>>> Temp.xml)



If "%Publisher%"=="*" (echo.    ^<Product ID="Publisher2019Retail"^>>> Temp.xml & (@If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)) & (@If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)) & (@If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)) & echo.    ^</Product^>>> Temp.xml)

If "%ProjectPro%"=="*" (echo.    ^<Product ID="ProjectPro2019Retail"^>>> Temp.xml & (@If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)) & (@If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)) & (@If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)) & echo.    ^</Product^>>> Temp.xml)

If "%VisioPro%"=="*" (echo.    ^<Product ID="VisioPro2019Retail"^>>> Temp.xml & (@If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)) & (@If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)) & (@If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)) & echo.    ^</Product^>>> Temp.xml)





goto start


:start
mode 135,15
cls


If "%Office%"=="*" (If "%ProjectPro%"=="*" (echo �� ��ࠫ� 祪���� Office � ProjectPro, �� ProjectPro �⠢���� ⮫쪮 � ०��� Separately, �.� �� �� ����� ���� ����祭 � Office. & Set konflikt=*))
If "%Office%"=="*" (If "%VisioPro%"=="*" (echo �� ��ࠫ� 祪���� Office � VisioPro, �� VisioPro �⠢���� ⮫쪮 � ०��� Separately, �.� �� �� ����� ���� ����祭 � Office. & Set konflikt=*))

If "%konflikt%"=="*" (echo. & echo ���⠢��� �⤥�쭮 �� Office? [y / n] & Set /p konflikt= :)
If /i "%konflikt%"=="y" (goto Office_konflikt)





If "%Separately%"=="*" (@If "%OneNote%"=="*" (echo. & echo �� ��ࠫ� 祪���� Separately � OneNote, �� OneNote �⠢���� ⮫쪮 � ०��� Office. & Set konflikt=*))
If "%Separately%"=="*" (@If "%Lync%"=="*" (echo. & echo �� ��ࠫ� 祪���� Separately � Lync, �� Lync �⠢���� ⮫쪮 � ०��� Office. & Set konflikt=*))
If "%Separately%"=="*" (@If "%Teams%"=="*" (echo. & echo �� ��ࠫ� 祪���� Separately � Teams, �� Teams �⠢���� ⮫쪮 � ०��� Office. & Set konflikt=*))
If "%Separately%"=="*" (@If "%Groove%"=="*" (echo. & echo �� ��ࠫ� 祪���� Separately � Groove, �� Groove �⠢���� ⮫쪮 � ०��� Office. & Set konflikt=*))

If "%konflikt%"=="*" (echo. & echo ���⠢��� ��� ���� ����� Office? [y / n] & Set /p konflikt= :)
If /i "%konflikt%"=="y" (goto Separately_konflikt)





echo.  ^</Add^>>> Temp.xml
echo.^</Configuration^>>> Temp.xml
Timeout 1 >nul


:startWOs
mode 65,15
cls

:ifDown
If "%Download%"=="*" (goto Download) else (goto Inst_Without_download)
:Download
goto Open_progresbar
:Open_progresbar_ok
cls
echo.
Echo ���樨஢��� ����㧪� ����室���� 䠩���.
Echo �� �����⨨ ������� ���� ����㧪� �㤥� ��ࢠ��.
Echo ����㧪� ����� ������ �����஥ �६�.
echo.
echo Loading...
start Temp.bat
setup.exe /download %xmlFileName%
powershell echo `a
cls
echo.
Echo ���樨஢��� ����㧪� ����室���� 䠩���.
Echo �� �����⨨ ������� ���� ����㧪� �㤥� ��ࢠ��.
Echo ����㧪� ����� ������ �����஥ �६�.
echo.
Echo Loading completed
goto ifInst

:: �᫨ ᪠稢���� �� �뫮 ��࠭�, ����� �ࠧ� ��⠭�������� ���
:Inst_Without_download
echo.
Echo ���樨஢��� ��⠭����. ������ �� ���� ����� �������.
Echo Installation initiated. Now this window can be closed.
echo.
Echo Installation...
echo.
setup.exe /configure %xmlFileName%
powershell echo `a
cls
echo.
Echo ���樨஢��� ��⠭����. ������ �� ���� ����� �������.
Echo Installation initiated. Now this window can be closed.
echo.
Echo Installation completed
echo.
pause
exit

:ifInst
If "%Install%"=="*" (goto Install) else (goto Without_Install)
:Install
cls
echo.
Echo ���樨஢��� ����㧪� ����室���� 䠩���.
Echo �� �����⨨ ������� ���� ����㧪�/��⠭���� �㤥� ��ࢠ��.
Echo ����㧪� ����� ������ �����஥ �६�.
echo.
Echo Loading completed
echo.
echo.
Echo ���樨஢��� ��⠭����. ������ �� ���� ����� �������.
Echo Installation initiated. Now this window can be closed.
echo.
Echo Installation...
setup.exe /configure %xmlFileName%
cls
echo.
Echo ���樨஢��� ����㧪� ����室���� 䠩���.
Echo �� �����⨨ ������� ���� ����㧪�/��⠭���� �㤥� ��ࢠ��.
Echo ����㧪� ����� ������ �����஥ �६�.
echo.
Echo Loading completed
echo.
echo.
Echo ���樨஢��� ��⠭����. ������ �� ���� ����� �������.
Echo Installation initiated. Now this window can be closed.
echo.
Echo Installation completed
powershell echo `a


:Without_Install
pause
exit








:Office_konflikt
echo.^<Product ID="ProjectPro2019Retail"^>>> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)
echo.    ^</Product^>>> Temp.xml

echo.^<Product ID="VisioPro2019Retail"^>>> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)
echo.    ^</Product^>>> Temp.xml

Set VisioPro=
Set ProjectPro=
Set konflikt=
goto start





:Separately_konflikt
echo.    ^<Product ID="O365ProPlusRetail"^> >> Temp.xml
If "%ru%"=="*" (echo.      ^<Language ID="ru-RU" /^>>> Temp.xml)
If "%eng%"=="*" (echo.      ^<Language ID="en-us" /^>>> Temp.xml)
If "%other%"=="*" (echo.      ^<Language ID="%LangAdd%" /^>>> Temp.xml)

If "%OneNote%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="OneNote" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="OneNote" /^>>> Temp.xml)
If "%Lync%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="Lync" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="Lync" /^>>> Temp.xml)
If "%Teams%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="Teams" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="Teams" /^>>> Temp.xml)
If "%Groove%"=="*" (echo.      ^<!-- ^<ExcludeApp ID="Groove" /^> --^>>> Temp.xml) else (echo.      ^<ExcludeApp ID="Groove" /^>>> Temp.xml)

echo.      ^<ExcludeApp ID="Word" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="Excel" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="PowerPoint" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="Access" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="Outlook" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="Publisher" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="Project" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="Visio" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="Skype" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="Skypeforbusiness" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="OneDriveforbusiness" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="InfoPath" /^>>> Temp.xml
echo.      ^<ExcludeApp ID="SharePointDesigner" /^>>> Temp.xml
echo.    ^</Product^> >> Temp.xml

Set OneNote=
Set Lync=
Set Teams=
Set Groove=
Set konflikt=
goto start







:Open_progresbar
echo.@echo off> Temp.bat
echo.>> Temp.bat
echo.Title Downloading status>> Temp.bat
echo.color 0A>> Temp.bat
echo.mode 23,10>> Temp.bat
echo.echo.>> Temp.bat
echo.Timeout 1 ^>nul >> Temp.bat
echo.:do_intro>> Temp.bat
echo.^<nul set /p strTemp=�>> Temp.bat
echo.ping -w 100 -n 1 127.0.0.1 ^> NUL>> Temp.bat
echo.set /a counter+=1 >> Temp.bat
echo.if %%counter%% geq 23 (goto next_intro) else (goto do_intro)>> Temp.bat
echo.:next_intro>> Temp.bat
echo.cls>> Temp.bat
echo.Set size=0 ^& Set b=0^& Set Mb=0 ^& Set Gb_DIV=0^& Set Gb_MOD=0 ^& Set perc=�����������������������>> Temp.bat
echo.set /a counter=0 >> Temp.bat
echo.SetLocal enabledelayedexpansion>> Temp.bat
echo.If not exist "%%cd%%\Office" md "%%cd%%"\Office>> Temp.bat
:: & echo.0123456789>"%cd%"\Office\Temp.txt
:: �⮡� ���� �訡�� "���������騩 ���࠭�" ����� � ����� Office ��� 䠩���
echo.Set D=%%cd%%\Office>> Temp.bat
echo.>> Temp.bat
echo.:do>> Temp.bat
echo.cls>> Temp.bat
echo.echo %%perc%%>> Temp.bat
echo.echo ----------------------->> Temp.bat
echo.echo   %%Gb_DIV%%,%%Gb_MOD:~0,2%% Gb or %%Mb%% Mb>> Temp.bat
echo.echo ----------------------->> Temp.bat
echo.echo byte     : !size!>> Temp.bat
echo.echo byte-1ch : %%b%%>> Temp.bat
echo.echo Mb   	 : %%Mb%%>> Temp.bat
echo.echo Gb_DIV   : %%Gb_DIV%%>> Temp.bat
echo.echo Gb_MOD   : %%Gb_MOD%%>> Temp.bat
::For /F "tokens=1-3" %%a IN ('Dir "%D%" /-C/S/A:-D') Do Set DirSize=!n2!& Set n2=%%c
::for /f "tokens=*" %%x in ('dir /s /a /b %1') do set /a size+=%%~zx
::echo.For /F "tokens=1-3" %%%%a IN ('Dir "%%D%%" /-C/S/A:-D') Do Set DirSize=!n2!^& Set n2=%%%%c>> Temp.bat
echo.for /f "tokens=3,5" %%%%a in ('dir /a /s /w /-c "%%D%%"^^^| findstr /b /l /c:"  "') do if "%%%%b"=="" set size=%%%%a^>nul>> Temp.bat
::Set DirSize=2147483647
echo.Set /a b=%%size:~0,-1%% >> Temp.bat
echo.Set /a Mb=(%%b%%">>"20)*10>> Temp.bat
echo.Set /a Gb_DIV=%%Mb%%">>"10 >> Temp.bat
echo.Set /a Gb_MOD=%%Mb%%%%%%1024 >> Temp.bat
echo.If %%Gb_MOD%% LEQ 99 Set Gb_MOD=0%%Gb_MOD%%>> Temp.bat
::
::start echo off ^& echo ------------------- ^& color 0A ^& mode 35,5 ^& echo. ^& echo %Gb:~0,1%,%Gb:~1% Gb or %Mb%0 Mb ^& ping -n 2 localhost >nul ^& exit
echo.::>> Temp.bat
echo.   if %%Mb%% GEQ 0 Set perc=�����������������������>> Temp.bat
echo. if %%Mb%% GEQ 272 Set perc=�����������������������>> Temp.bat
echo. if %%Mb%% GEQ 544 Set perc=�����������������������>> Temp.bat
echo. if %%Mb%% GEQ 816 Set perc=�����������������������>> Temp.bat
echo.if %%Mb%% GEQ 1088 Set perc=�����������������������>> Temp.bat
echo.if %%Mb%% GEQ 1360 Set perc=�����������������������>> Temp.bat
echo.if %%Mb%% GEQ 1632 Set perc=�����������������������>> Temp.bat
echo.if %%Mb%% GEQ 1904 Set perc=�����������������������>> Temp.bat
echo.if %%Mb%% GEQ 2176 Set perc=�����������������������>> Temp.bat
echo.if %%Mb%% GEQ 2448 Set perc=�����������������������>> Temp.bat
echo.if %%Mb%% GEQ 2720 Set perc=�����������������������>> Temp.bat
echo.set /a counter+=1 >> Temp.bat
echo.If %%counter%% GEQ 2 timeout 8 ^>nul >> Temp.bat
::echo.timeout 2 ^>nul >> Temp.bat
echo.:while >> Temp.bat
echo.if %%Mb%% geq 2720 (goto next) else (goto do) >> Temp.bat
echo.:next >> Temp.bat
::echo.timeout 2 ^>nul >> Temp.bat
echo.echo.>> Temp.bat
echo.>> Temp.bat
echo.echo Completed >> Temp.bat
echo.timeout 2 ^>nul >> Temp.bat
::echo.pause ^>nul >> Temp.bat
echo.del "%%~f0" ^& exit >> Temp.bat


goto Open_progresbar_ok















