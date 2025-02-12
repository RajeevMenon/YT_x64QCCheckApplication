Imports TML_Library

#Region "SPARK-DOMAIN"
'Public NotInheritable Class Link
'    Public Const PlanID As String = "5Q00"
'    Public Const Network As String = "DOMAIN"
'    Public Const ServerAddress As String = "cpc475-tmlpc03.bh.me.ykgw.net" '"10.25.54.130" '
'    Public Const PortNo As String = "4600"
'    Public Const FE0_File As String = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=$$$\FE0_YMA.mdb" 'Replace, $$$ with Application.StartupPath or As needed
'    Public Const Mysql_YGSP_ConStr As String = "Server=" & ServerAddress & ";Port=" & PortNo & ";Database=LIVE_YGSP;Uid=PRDSCH;Pwd=LMLINE;CharSet=utf8;"
'    Public Const Mysql_QA_ConStr As String = "Server=" & ServerAddress & ";Port=" & PortNo & ";Database=LIVE_QA;Uid=PRDSCH;Pwd=LMLINE;CharSet=utf8;"
'    Public Const ServerDB As String = "LIVE_YGSP"
'    Public Const FactoryNetwork As Initialize.FACTORY = Initialize.FACTORY._YKSA_M_NET
'    Public Const FactoryQA As String = "A.Ahmed"
'    Public Const Factory As String = "2"
'    Public Const DocsStore As String = "\\" & ServerAddress & "\DocsStore"
'    Public Const MTC_BMP_Path As String = "\\" & ServerAddress & "\MatCertShare"
'    Public Const CompleteDocumentFolder As String = DocsStore & "\Production Complete Documents"
'    Public Const QIC_Template_Linear As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicLinear.rpt"
'    Public Const QIC_Template_Sqrt As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicSquareRoot.rpt"
'    Public Const QIC_Template_Linear9P As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicLinear9P.rpt"
'    Public Const QIC_Template_Recalibration As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\reCalibration.rpt"
'    Public Const LTC_Template_N2 As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest.rpt"
'    Public Const LTC_Template_H2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_70MPa.rpt"
'    Public Const MNFTC_Template_N2 As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\MNF_PressureTest.rpt"
'    Public Const MNFTC_Template_H2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_70MPa.rpt"
'    Public Const MTC_Template As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\MillCertificate_v01.rpt"
'    Public Const DEVINFO_Template As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\Device_Info.xls"
'    Public Const Server As String = "\\" & ServerAddress & "\"
'    Public Const CrystallTemplateFolder As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\"
'    Public Const OrderTagLocation As String = "\\" & ServerAddress & "\DocsStore\Production Release Documents\OrderTag\"
'    Public Const PmiCertificatePath As String = "\\" & ServerAddress & "\DocsStore\Production Complete Documents\PMI_Certificates\"
'    Public Const CapUsageServer_1 As String = "10.7.15.25"
'    Public Const CapUsageServer_2 As String = "10.7.15.140"
'    Public Const CapUsagePort As Long = 2035
'End Class
#End Region

#Region "SPARK-FACTORY"
Public NotInheritable Class Link
    Public Const PlanID As String = "5Q00"
    Public Const Network As String = "FACTORY"
    Public Const ServerAddress As String = "192.168.151.12"
    Public Const PortNo As String = "4600"
    Public Const FE0_File As String = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=$$$\FE0_YMA.mdb" 'Replace, $$$ with Application.StartupPath or As needed
    Public Const Mysql_YGSP_ConStr As String = "Server=192.168.151.12;Port=4600;Database=LIVE_YGSP;Uid=PRDSCH;Pwd=LMLINE;CharSet=utf8;"
    Public Const Mysql_QA_ConStr As String = "Server=192.168.151.12;Port=4600;Database=LIVE_QA;Uid=PRDSCH;Pwd=LMLINE;CharSet=utf8;"
    Public Const ServerDB As String = "LIVE_YGSP"
    Public Const FactoryNetwork As Initialize.FACTORY = Initialize.FACTORY._YKSA_F_NET
    Public Const FactoryQA As String = "A.Ahmed"
    Public Const Factory As String = "4"
    Public Const DocsStore As String = "\\" & ServerAddress & "\DocsStore"
    Public Const MTC_BMP_Path As String = "\\" & ServerAddress & "\MatCertShare"
    Public Const CompleteDocumentFolder As String = DocsStore & "\Production Complete Documents"
    Public Const QIC_Template_Linear As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicLinear.rpt"
    Public Const QIC_Template_Sqrt As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicSquareRoot.rpt"
    Public Const QIC_Template_Linear9P As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicLinear9P.rpt"
    Public Const QIC_Template_Recalibration As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\reCalibration.rpt"
    Public Const LTC_Template_N2 As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest.rpt"
    Public Const LTC_Template_H2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_70MPa.rpt"
    Public Const MNFTC_Template_N2 As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\MNF_PressureTest.rpt"
    Public Const MNFTC_Template_H2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_70MPa.rpt"
    Public Const MTC_Template As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\MillCertificate_v01.rpt"
    Public Const DEVINFO_Template As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\Device_Info.xls"
    Public Const Server As String = "\\" & ServerAddress & "\"
    Public Const CrystallTemplateFolder As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\"
    Public Const OrderTagLocation As String = "\\" & ServerAddress & "\DocsStore\Production Release Documents\OrderTag\"
    Public Const PmiCertificatePath As String = "\\" & ServerAddress & "\DocsStore\Production Complete Documents\PMI_Certificates\"
    Public Const CapUsageServer_1 As String = "10.7.15.25"
    Public Const CapUsageServer_2 As String = "10.7.15.140"
    Public Const CapUsagePort As Long = 2035
End Class
#End Region

#Region "ICAD-DOMAIN"
'Public NotInheritable Class Link
'    Public Const PlanID As String = "6Z00"
'    Public Const Network As String = "DOMAIN"
'    Public Const ServerAddress As String = "cpc475-tmlpc04.bh.me.ykgw.net" '"10.25.216.12" '
'    Public Const PortNo As String = "4600"
'    Public Const FE0_File As String = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=$$$\FE0_YMA.mdb" 'Replace, $$$ with Application.StartupPath or As needed
'    Public Const Mysql_YGSP_ConStr As String = "Server=" & ServerAddress & ";Port=" & PortNo & ";Database=LIVE_YGSP;Uid=PRDSCH;Pwd=LMLINE;CharSet=utf8;"
'    Public Const Mysql_QA_ConStr As String = "Server=" & ServerAddress & ";Port=" & PortNo & ";Database=LIVE_QA;Uid=PRDSCH;Pwd=LMLINE;CharSet=utf8;"
'    Public Const ServerDB As String = "LIVE_YGSP"
'    Public Const FactoryNetwork As Initialize.FACTORY = Initialize.FACTORY._YUAE_M_NET
'    Public Const FactoryQA As String = "K.Gim"
'    Public Const Factory As String = "1"
'    Public Const DocsStore As String = "\\" & ServerAddress & "\DocsStore"
'    Public Const MTC_YhqFiles As String = "\\" & ServerAddress & "\MatCertShare\"
'    Public Const MTC_Settings As String = "\\" & ServerAddress & "\DocsStore\SOFTWARES\LIVE SOFTWARES\COMMON\MTC_DATA\"
'    Public Const CompleteDocumentFolder As String = DocsStore & "\Production Complete Documents"
'    Public Const QIC_Template_Linear As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicLinear.rpt"
'    Public Const QIC_Template_Sqrt As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicSquareRoot.rpt"
'    Public Const QIC_Template_Linear9P As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicLinear9P.rpt"
'    Public Const QIC_Template_Recalibration As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\reCalibration.rpt"
'    Public Const LTC_Template_N2 As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest.rpt"
'    Public Const LTC_Template_H2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_70MPa.rpt"
'    Public Const LTC_Template_DPH2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_DP70MPa.rpt"
'    Public Const MNFTC_Template_N2 As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\MNF_PressureTest.rpt"
'    Public Const MNFTC_Template_H2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_70MPa.rpt"
'    Public Const MTC_Template As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\MillCertificate_v01.rpt"
'    Public Const DEVINFO_Template As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\Device_Info.xls"
'    Public Const Server As String = "\\" & ServerAddress & "\"
'    Public Const CrystallTemplateFolder As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\"
'    Public Const OrderTagLocation As String = "\\" & ServerAddress & "\DocsStore\Production Release Documents\OrderTag\"
'    Public Const PmiCertificatePath As String = "\\" & ServerAddress & "\DocsStore\Production Complete Documents\PMI_Certificates\"
'    Public Const CapUsageServer_1 As String = "10.7.15.25"
'    Public Const CapUsageServer_2 As String = "10.7.15.140"
'    Public Const CapUsagePort As Long = 2035
'End Class
#End Region

#Region "ICAD-FACTORY"
'Public NotInheritable Class Link
'    Public Const PlanID As String = "6Z00"
'    Public Const Network As String = "FACTORY"
'    Public Const ServerAddress As String = "192.168.151.12"
'    Public Const PortNo As String = "4600"
'    Public Const FE0_File As String = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=$$$\FE0_YMA.mdb" 'Replace, $$$ with Application.StartupPath or As needed
'    Public Const Mysql_YGSP_ConStr As String = "Server=" & ServerAddress & ";Port=4600;Database=LIVE_YGSP;Uid=PRDSCH;Pwd=LMLINE;CharSet=utf8;"
'    Public Const Mysql_QA_ConStr As String = "Server=" & ServerAddress & ";Port=4600;Database=LIVE_QA;Uid=PRDSCH;Pwd=LMLINE;CharSet=utf8;"
'    Public Const ServerDB As String = "LIVE_YGSP"
'    Public Const FactoryNetwork As Initialize.FACTORY = Initialize.FACTORY._YUAE_F_NET
'    Public Const FactoryQA As String = "K.Gim"
'    Public Const Factory As String = "3"
'    Public Const DocsStore As String = "\\" & ServerAddress & "\DocsStore"
'    Public Const MTC_YhqFiles As String = "\\" & ServerAddress & "\MatCertShare\"
'    Public Const MTC_Settings As String = "\\" & ServerAddress & "\DocsStore\SOFTWARES\LIVE SOFTWARES\COMMON\MTC_DATA\"
'    Public Const CompleteDocumentFolder As String = DocsStore & "\Production Complete Documents"
'    Public Const QIC_Template_Linear As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicLinear.rpt"
'    Public Const QIC_Template_Sqrt As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicSquareRoot.rpt"
'    Public Const QIC_Template_Linear9P As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicLinear9P.rpt"
'    Public Const QIC_Template_Recalibration As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\reCalibration.rpt"
'    Public Const LTC_Template_N2 As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest.rpt"
'    Public Const LTC_Template_H2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_70MPa.rpt"
'    Public Const LTC_Template_DPH2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_DP70MPa.rpt"
'    Public Const MNFTC_Template_N2 As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\MNF_PressureTest.rpt"
'    Public Const MNFTC_Template_H2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_70MPa.rpt"
'    Public Const MTC_Template As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\MillCertificate_v01.rpt"
'    Public Const DEVINFO_Template As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\Device_Info.xls"
'    Public Const Server As String = "\\" & ServerAddress & "\"
'    Public Const CrystallTemplateFolder As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\"
'    Public Const OrderTagLocation As String = "\\" & ServerAddress & "\DocsStore\Production Release Documents\OrderTag\"
'    Public Const PmiCertificatePath As String = "\\" & ServerAddress & "\DocsStore\Production Complete Documents\PMI_Certificates\"
'    Public Const CapUsageServer_1 As String = "10.7.15.25"
'    Public Const CapUsageServer_2 As String = "10.7.15.140"
'    Public Const CapUsagePort As Long = 2035
'End Class
#End Region

#Region "LOCAL-TEST"
'Public NotInheritable Class Link
'    Public Const PlanID As String = "6Z00"
'    Public Const Network As String = "LOCAL"
'    Public Const ServerAddress As String = "127.0.0.1"
'    Public Const FE0_File As String = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=$$$\FE0_YMA.mdb" 'Replace, $$$ with Application.StartupPath or As needed
'    Public Const Mysql_YGSP_ConStr As String = "Server=127.0.0.1;Port=4600;Database=LIVE_YGSP;Uid=PRDSCH;Pwd=LMLINE;CharSet=utf8;"
'    Public Const Mysql_QA_ConStr As String = "Server=127.0.0.1;Port=4600;Database=LIVE_QA;Uid=PRDSCH;Pwd=LMLINE;CharSet=utf8;"
'    Public Const ServerDB As String = "LIVE_YGSP"
'    Public Const Factory As String = "1"
'    Public Const FactoryNetwork As Initialize.FACTORY = Initialize.FACTORY._YUAE_M_NET
'    Public Const FactoryQA As String = "K.Gim"
'    Public Const DocsStore As String = "\\" & ServerAddress & "\DocsStore"
'    Public Const MTC_YhqFiles As String = "\\" & ServerAddress & "\MatCertShare\"
'    Public Const MTC_Settings As String = "\\" & ServerAddress & "\DocsStore\SOFTWARES\LIVE SOFTWARES\COMMON\MTC_DATA\"
'    Public Const MTC_BMP_Path As String = "\\" & ServerAddress & "\MatCertShare"
'    Public Const CompleteDocumentFolder As String = DocsStore & "\Production Complete Documents"
'    Public Const QIC_Template_Linear As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicLinear.rpt"
'    Public Const QIC_Template_Sqrt As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicSquareRoot.rpt"
'    Public Const QIC_Template_Linear9P As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\xQicLinear9P.rpt"
'    Public Const QIC_Template_Recalibration As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\reCalibration.rpt"
'    Public Const LTC_Template_N2 As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest.rpt"
'    Public Const LTC_Template_H2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_70MPa.rpt"
'    Public Const LTC_Template_DPH2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_DP70MPa.rpt"
'    Public Const MNFTC_Template_N2 As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\MNF_PressureTest.rpt"
'    Public Const MNFTC_Template_H2O As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\PressureTest_70MPa.rpt"
'    Public Const MTC_Template As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\MillCertificate_v01.rpt"
'    Public Const DEVINFO_Template As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\Device_Info.xls"
'    Public Const Server As String = "\\" & ServerAddress & "\"
'    Public Const CrystallTemplateFolder As String = "\\" & ServerAddress & "\DocsStore\Common Documents\Document Templates\"
'    Public Const OrderTagLocation As String = "\\" & ServerAddress & "\DocsStore\Production Release Documents\OrderTag\"
'    Public Const PmiCertificatePath As String = "\\" & ServerAddress & "\DocsStore\Production Complete Documents\PMI_Certificates\"
'    Public Const CapUsageServer_1 As String = "10.7.15.25"
'    Public Const CapUsageServer_2 As String = "10.7.15.140"
'    Public Const CapUsagePort As Long = 2035
'End Class
#End Region
