@ECHO OFF
::********************************************START  SCRIPT INPUTS********************************************
:: All password inputs (*Pwd) can be left blank for security reasons; the script will prompt you for input.
:: ALL other inputs are REQUIRED at this point.
:: Note: if you're using the same input (e.g. accounts) for different services, you can refer to previously-defined input
:: e.g. SET MySitesAppPoolAcct=%SSPAppPoolAcct%
:: *************************************************************************************************************
SET                       DBServer=
SET          CentralAdminContentDB=%COMPUTERNAME%-MOSS_CentralAdmin_Content 
SET                       ConfigDB=%COMPUTERNAME%-MOSS_Config
SET              SSPAdminContentDB=%COMPUTERNAME%-MOSS_SSPAdmin_Content
SET                   SSPContentDB=%COMPUTERNAME%-MOSS_SSP_Content
SET                    SSPSearchDB=%COMPUTERNAME%-MOSS_SSP_Search
SET               MySitesContentDB=%COMPUTERNAME%-MOSS_MySites_Content
SET                    WSSSearchDB=%COMPUTERNAME%-MOSS_WSS_Search
SET                PortalContentDB=%COMPUTERNAME%-MOSS_Portal_Content
SET                       FarmAcct=%USERDOMAIN%\MOSS_FarmAccount
SET                    FarmAcctPWD=MOSS_FarmAccount
SET                  FarmAcctEmail=moss@%USERDOMAIN%.ca
SET                  WSSSearchAcct=%USERDOMAIN%\WSS_SearchService
SET               WSSSearchAcctPWD=WSS_SearchService
SET               SearchAccessAcct=%USERDOMAIN%\MOSS_SearchAccess
SET            SearchAccessAcctPWD=MOSS_SearchAccess
SET                 MOSSSearchAcct=%USERDOMAIN%\MOSS_SearchService
SET              MOSSSearchAcctPWD=MOSS_SearchService
SET                 SSPAppPoolAcct=%USERDOMAIN%\MOSS_SSPAdminAppPool
SET              SSPAppPoolAcctPWD=MOSS_SSPAdminAppPool
SET             MySitesAppPoolAcct=%USERDOMAIN%\MOSS_MySitesAppPool
SET          MySitesAppPoolAcctPWD=MOSS_MySitesAppPool
SET              PortalAppPoolAcct=%USERDOMAIN%\MOSS_IntranetAppPool
SET           PortalAppPoolAcctPWD=MOSS_IntranetAppPool
SET                InstallLocation=%PROGRAMFILES%\Microsoft Office Servers
SET                        DataDir=%PROGRAMFILES%\Microsoft Office Servers\12.0\Data
SET               CentralAdminPort=2007
SET                        SSPDesc=SSP1
SET                         SSPURL=http://%COMPUTERNAME%
SET                        SSPPort=8081
SET                    MySitesDesc=My_Sites
SET                     MySitesURL=http://%COMPUTERNAME%
SET                    MySitesPORT=8082
SET                     PortalDesc=Portal
SET                      PortalURL=http://%COMPUTERNAME%
SET                     PortalPort=80
SET                 PortalTemplate=SPSPORTAL
SET               SSPIndexLocation=%DataDir%\Office Server\Applications
SET                 SSPIndexServer=%COMPUTERNAME%
SET               WSSIndexLocation=
SET              MOSSIndexLocation=%DataDir%\Office Server\Applications
::********************************************END  SCRIPT INPUTS**********************************************