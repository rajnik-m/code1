Namespace Access

  Public Class AccessControl

    Private mvEnv As CDBEnvironment

    Public Sub New(pEnv As CDBEnvironment)
      mvEnv = pEnv
    End Sub

    Public Sub Initialise()
      Dim vItemFields As New CDBFields
      Dim vAreaFields As New CDBFields

      'Delete Access Control Items
      mvEnv.Connection.DeleteAllRecords("access_control_items")

      'Delete Access Control Area
      mvEnv.Connection.DeleteAllRecords("access_control_areas")

      'Delete any existing MAIN group
      Dim vFields As New CDBFields(New CDBField("access_control_group", "MAIN"))
      mvEnv.Connection.DeleteRecords("access_control_groups", vFields, False)

      'Now add the default group
      vFields.Add("access_control_group_desc", CDBField.FieldTypes.cftCharacter, "Default Access Control Group")
      vFields.AddAmendedOnBy(mvEnv.User.UserID, TodaysDate)
      mvEnv.Connection.InsertRecord("access_control_groups", vFields)

      'Now set up the Access Control Areas
      vFields = New CDBFields
      vFields.AddAmendedOnBy(mvEnv.User.UserID, TodaysDate)
      vFields.Add("access_control_area")
      vFields.Add("access_control_area_desc")
      vFields.Add("sequence_number")

      'Smart Client
      AddAccessControlArea(mvEnv, vFields, "SC,Smart Client,60000")
      AddAccessControlArea(mvEnv, vFields, "SCFL,File Menu,61000")
      AddAccessControlArea(mvEnv, vFields, "SCVM,View Menu,62000")
      AddAccessControlArea(mvEnv, vFields, "SCQM,Query Menu,62500")
      AddAccessControlArea(mvEnv, vFields, "SCFM,Find Menu,63000")
      AddAccessControlArea(mvEnv, vFields, "SCTM,Tools Menu,64000")
      AddAccessControlArea(mvEnv, vFields, "SCSM,System Menu,65000")
      AddAccessControlArea(mvEnv, vFields, "SCAM,Administration Menu,66000")
      AddAccessControlArea(mvEnv, vFields, "SCFLNE,New,67000")
      AddAccessControlArea(mvEnv, vFields, "SCCP,Campaign PopUp Menu,68000")
      AddAccessControlArea(mvEnv, vFields, "SCLM,List Manager,69000")
      AddAccessControlArea(mvEnv, vFields, "SCSS,Selection Set PopUp Menu,70000")
      AddAccessControlArea(mvEnv, vFields, "SCBM,Browser PopUp Menu,71000")
      AddAccessControlArea(mvEnv, vFields, "SCPR,Preferences,72000")
      AddAccessControlArea(mvEnv, vFields, "SCFP,Financial PopUp Menu,73000")
      AddAccessControlArea(mvEnv, vFields, "SCDA,Dashboard,74000")
      AddAccessControlArea(mvEnv, vFields, "SCDP,Dashboard PopUp Menu,75000")
      AddAccessControlArea(mvEnv, vFields, "SCMP,Mailing PopUp Menu,76000")
      AddAccessControlArea(mvEnv, vFields, "SCCU,Customise PopUp Menu,77000")
      AddAccessControlArea(mvEnv, vFields, "CDDP,Document PopUp Menu,78000")
      AddAccessControlArea(mvEnv, vFields, "SCEM,Event Maintenance,79000")
      AddAccessControlArea(mvEnv, vFields, "SCCO,Contacts,79030")
      AddAccessControlArea(mvEnv, vFields, "SCGE,General,79050")

      AddAccessControlArea(mvEnv, vFields, "SCSMBA,Banks,79100")
      AddAccessControlArea(mvEnv, vFields, "SCSMBM,Batch Management,79200")
      AddAccessControlArea(mvEnv, vFields, "SCSMCA,CAF,79300")
      AddAccessControlArea(mvEnv, vFields, "SCSMCP,CPD,79450")
      AddAccessControlArea(mvEnv, vFields, "SCSMCC,Credit Cards,79500")
      AddAccessControlArea(mvEnv, vFields, "SCSMCS,Credit Sales,79600")
      AddAccessControlArea(mvEnv, vFields, "SCSMFD,Data Entry Application,79700")
      AddAccessControlArea(mvEnv, vFields, "SCSMDD,De-Duplication,79800")
      AddAccessControlArea(mvEnv, vFields, "SCSMDB,Direct Debits,79900")
      AddAccessControlArea(mvEnv, vFields, "SCSMDI,Distribution Boxes,80000")
      AddAccessControlArea(mvEnv, vFields, "SCSMDR,Distribution Reports,80050")
      AddAccessControlArea(mvEnv, vFields, "SCSMDU,Dutch Payment Processing,80100")
      AddAccessControlArea(mvEnv, vFields, "SCSMCM,E-Marketing,80150")
      AddAccessControlArea(mvEnv, vFields, "SCSMEV,Events,80150")
      AddAccessControlArea(mvEnv, vFields, "SCSMEX,Exams,80160")
      AddAccessControlArea(mvEnv, vFields, "SCSMGA,Gift Aid Declarations,80200")
      AddAccessControlArea(mvEnv, vFields, "SCSMGS,Gift Aid Sponsorship,80210")
      AddAccessControlArea(mvEnv, vFields, "SCSMGI,Irish Gift Aid,80220")
      AddAccessControlArea(mvEnv, vFields, "SCSMIN,Incentives,80300")
      AddAccessControlArea(mvEnv, vFields, "SCSMMA,Mailings,80400")
      AddAccessControlArea(mvEnv, vFields, "SCSMMK,Marketing,80500")
      AddAccessControlArea(mvEnv, vFields, "SCSMME,Membership,80600")
      AddAccessControlArea(mvEnv, vFields, "SCSMRP,Membership Reports,80610")
      AddAccessControlArea(mvEnv, vFields, "SCSMST,Membership Statistics,80620")
      AddAccessControlArea(mvEnv, vFields, "SCSMNC,Nominal Codes,80700")
      AddAccessControlArea(mvEnv, vFields, "SCSMPI,Paying In Slips,80800")
      AddAccessControlArea(mvEnv, vFields, "SCSMPP,Payment Plans,80900")
      AddAccessControlArea(mvEnv, vFields, "SCSMPG,Payroll Giving,81000")
      AddAccessControlArea(mvEnv, vFields, "SCSMPR,Purchase Orders,81100")
      AddAccessControlArea(mvEnv, vFields, "SCSMPD,Products,81200")
      AddAccessControlArea(mvEnv, vFields, "SCSMSO,Standing Orders,81300")
      AddAccessControlArea(mvEnv, vFields, "SCSMSK,Stock,81400")
      AddAccessControlArea(mvEnv, vFields, "SCSMUP,Updates,81500")
      AddAccessControlArea(mvEnv, vFields, "CDEV,Events,81600")
      AddAccessControlArea(mvEnv, vFields, "CDFP,Financial PopUp Menu,81700")
      AddAccessControlArea(mvEnv, vFields, "CDAP,Account PopUp Menu,81900")
      AddAccessControlArea(mvEnv, vFields, "SCGM,Mailings,82000")
      AddAccessControlArea(mvEnv, vFields, "SCXA,Exam Maintenance,82100")
      AddAccessControlArea(mvEnv, vFields, "SCXACO,Courses,82200")
      AddAccessControlArea(mvEnv, vFields, "SCXACE,Centres,82300")
      AddAccessControlArea(mvEnv, vFields, "SCXAPE,Personnel,82400")
      AddAccessControlArea(mvEnv, vFields, "SCXASE,Sessions,82500")
      AddAccessControlArea(mvEnv, vFields, "SCXAXE,Exemptions,82600")
      AddAccessControlArea(mvEnv, vFields, "SCXAPU,PopUp Menu,82700")
      AddAccessControlArea(mvEnv, vFields, "SCPO,Purchase Order PopUp Menu,82800")

      'Add Access Control Items
      vFields = New CDBFields
      vFields.AddAmendedOnBy(mvEnv.User.Logname)
      vFields.Add("access_control_item")
      vFields.Add("access_control_item_desc")
      vFields.Add("access_control_area")
      vFields.Add("sequence_number")
      vFields.Add("access_level")
      vFields.Add("system_module")

      'Smart Client - File Menu
      AddAccessControlItem(mvEnv, vFields, "SCLMPR,Preferences,SCFL,1000,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCLMLW,Log WEB Service Calls,SCFL,1100,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCLMPS,Page Setup,SCFL,1200,U,CD")

      'Smart Client - File Menu - New
      AddAccessControlItem(mvEnv, vFields, "SCFLNC,Contact,SCFLNE,1000,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLC2,Contact2,SCFLNE,1005,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLC3,Contact3,SCFLNE,1010,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLC4,Contact4,SCFLNE,1015,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLC5,Contact5,SCFLNE,1020,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLNO,Organisation,SCFLNE,1100,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLO2,Organisation2,SCFLNE,1105,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLO3,Organisation3,SCFLNE,1110,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLO4,Organisation4,SCFLNE,1115,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLO5,Organisation5,SCFLNE,1120,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLND,Document,SCFLNE,1200,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLNA,Action,SCFLNE,1300,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLNB,Action Template,SCFLNE,1300,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLNT,Telephone Call,SCFLNE,1400,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLNS,Selection Set,SCFLNE,1400,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFLEV,Event,SCFLNE,1500,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCFLE2,Event2,SCFLNE,1505,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCFLE3,Event3,SCFLNE,1510,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCFLE4,Event4,SCFLNE,1515,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCFLE5,Event5,SCFLNE,1520,U,EV")

      'Smart Client - View Menu
      AddAccessControlItem(mvEnv, vFields, "SCVMTB,Tool Bar,SCVM,1000,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMNP,Navigation Panel,SCVM,1100,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMSB,Status Bar,SCVM,1200,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMHP,Header Panel,SCVM,1300,N,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMSP,Selection Panel,SCVM,1400,N,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMDB,Dashboard,SCVM,1500,N,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMMT,My Details,SCVM,1600,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMMO,My Organisation,SCVM,1700,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMMA,My Actions,SCVM,1800,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMMD,My Documents,SCVM,1900,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMMI,My InBox,SCVM,2000,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMMJ,My Journal,SCVM,2100,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCVMRE,Refresh,SCVM,2200,R,CD")

      'Smart Client - Find Menu
      AddAccessControlItem(mvEnv, vFields, "SCFMSE,Search Data,SCFM,1100,N,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMPF,Contact Finder,SCFM,1105,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMP2,Contact Finder2,SCFM,1110,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMP3,Contact Finder3,SCFM,1115,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMP4,Contact Finder4,SCFM,1120,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMP5,Contact Finder5,SCFM,1125,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMOF,Organisation Finder,SCFM,1200,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMO2,Organisation Finder2,SCFM,1205,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMO3,Organisation Finder3,SCFM,1210,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMO4,Organisation Finder4,SCFM,1215,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMO5,Organisation Finder5,SCFM,1220,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMDF,Document Finder,SCFM,1300,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMMF,Meeting Finder,SCFM,1600,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMAF,Action Finder,SCFM,1605,U,AC")
      AddAccessControlItem(mvEnv, vFields, "SCFMSS,Selection Sets,SCFM,1610,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMEF,Event Finder,SCFM,1650,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCFME2,Event Finder2,SCFM,1655,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCFME3,Event Finder3,SCFM,1660,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCFME4,Event Finder4,SCFM,1665,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCFME5,Event Finder5,SCFM,1670,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCFMBF,Member Finder,SCFM,1700,R,ME")
      AddAccessControlItem(mvEnv, vFields, "SCFMPP,Payment Plan Finder,SCFM,1800,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMCF,Covenant Finder,SCFM,1900,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMTF,Transaction Finder,SCFM,2000,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMSO,Standing Order Finder,SCFM,2100,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMDD,Direct Debit Finder,SCFM,2200,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMCC,CCCA Finder,SCFM,2300,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMGA,Gift Aid Declaration Finder,SCFM,2400,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMIF,Invoice/Credit Note Finder,SCFM,2500,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMLG,Legacy Finder,SCFM,2600,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMGY,Pre Tax Payroll Giving Finder,SCFM,2700,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMPG,Post Tax Payroll Giving Finder,SCFM,2800,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMPO,Purchase Order Finder,SCFM,2900,R,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMPC,Product Finder,SCFM,3000,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMCA,Campaign Finder,SCFM,3100,D,MA")
      AddAccessControlItem(mvEnv, vFields, "SCFMSD,Standard Document Finder,SCFM,3200,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMFP,Fundraising Payment Finder,SCFM,3300,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMSP,Service Product Finder,SCFM,3400,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCFMFR,Fundraising Request Finder,SCFM,3500,U,CD")

      'Samrt client - Query Menu
      AddAccessControlItem(mvEnv, vFields, "SCQMBE,Query By Example,SCQM,1105,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCQMPF,Query Contact,SCQM,1110,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCQMP2,Query Contact2,SCQM,1115,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCQMP3,Query Contact3,SCQM,1120,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCQMP4,Query Contact4,SCQM,1125,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCQMP5,Query Contact5,SCQM,1130,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCQMOF,Query Organisation,SCQM,1200,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCQMO2,Query Organisation2,SCQM,1205,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCQMO3,Query Organisation3,SCQM,1210,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCQMO4,Query Organisation4,SCQM,1215,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCQMO5,Query Organisation5,SCQM,1220,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCQMEF,Query Event,SCQM,1650,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCQME2,Query Event2,SCQM,1655,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCQME3,Query Event3,SCQM,1660,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCQME4,Query Event4,SCQM,1665,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCQME5,Query Event5,SCQM,1670,U,EV")

      'Smart Client - Tools Menu
      AddAccessControlItem(mvEnv, vFields, "SCTMTM,Table Maintenance,SCTM,1000,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCTMJS,Job Schedule,SCTM,1100,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCTMLM,List Manager,SCTM,1200,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCTMRM,Run Mailing,SCTM,1300,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCTMRR,Run Report,SCTM,1350,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCTMDD,Document Distributor,SCTM,1500,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCTMSM,Send EMail,SCTM,1600,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCTMPP,Postcode Proximity,SCTM,1700,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCTMEX,Explore,SCTM,1800,R,CD")
      AddAccessControlItem(mvEnv, vFields, "SCTMCU,Customise,SCTM,1900,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCTMCO,Close Open Batch,SCTM,2000,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCTMCM,Copy Event Pricing Matrix,SCTM,2100,D,CD")

      'Smart Client - System Menu
      '---------------------------------------------------------------------------------
      'Smart Client - System Menu - Banks
      AddAccessControlItem(mvEnv, vFields, "SCBALD,Load Bank Data,SCSMBA,1000,S,FM")

      'Smart Client - System Menu - Batch Management
      AddAccessControlItem(mvEnv, vFields, "SCSMVB,View Batch Details,SCSMBM,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMPB,Process Batches,SCSMBM,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMOS,Outstanding Batches Report,SCSMBM,1200,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMPC,Print Cheque List,SCSMBM,1300,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMPS,Print Paying In Slip,SCSMBM,1400,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMCB,Redo Cash Book Batch,SCSMBM,1500,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMPO,Post Batch,SCSMBM,1600,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMCJ,Create Journal Files,SCSMBM,1700,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMSR,Summary Report,SCSMBM,1800,U,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMDT,Detail Report,SCSMBM,1900,U,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMPU,Purge Old Batches,SCSMBM,2000,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMOB,Close Open Batch,SCSMBM,2100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCSMPZ,Purge Prize Draw Batches,SCSMBM,2200,S,FM")

      'Smart Client - System Menu - CAF
      AddAccessControlItem(mvEnv, vFields, "SCFCEP,SO/CCCA Expected Payments,SCSMCA,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFCPB,Voucher Claim,SCSMCA,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFCCS,Manual CAF Card Sales Claim,SCSMCA,1200,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFCPL,Load Payment Data,SCSMCA,1300,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFCPR,Reconcile Payment Data,SCSMCA,1400,S,FM")

      'Smart Client - System Menu - CheetahMail
      AddAccessControlItem(mvEnv, vFields, "SCSCMD,Process Meta Data,SCSMCM,1000,S,SC")
      AddAccessControlItem(mvEnv, vFields, "SCSCME,Process Event Data,SCSMCM,1100,S,SC")
      AddAccessControlItem(mvEnv, vFields, "SCSCMT,Process Mailing Totals,SCSMCM,1200,S,SC")
      AddAccessControlItem(mvEnv, vFields, "SCSBLK,Update Bulk Mailer Statistics,SCSMCM,1300,S,SC")

      'Smart Client - System Menu - Credit Cards
      AddAccessControlItem(mvEnv, vFields, "SCFCCB,Create CCCA Batches,SCSMCC,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFCCF,CCCA Claim File Creation,SCSMCC,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFCMC,Manual CCCA Claim,SCSMCC,1200,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFCSF,Card Sales Claim File Creation,SCSMCC,1300,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFCMS,Manual Card Sales Claim,SCSMCC,1400,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFCAR,Authorisations Report,SCSMCC,1500,S,FM")

      'Smart Client - System Menu - Credit Sales
      AddAccessControlItem(mvEnv, vFields, "SCFMTI,Transfer Invoices,SCSMCS,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMTC,Transfer Customers,SCSMCS,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFMSG,Statement Generation,SCSMCS,1200,s,FM")
      AddAccessControlItem(mvEnv, vFields, "CDCSNC,New Credit Customer,SCSMCS,1300,S,FM")
      AddAccessControlItem(mvEnv, vFields, "CDCSRS,Remove CreditStopCode,SCSMCS,1400,S,FM")

      'Smart Client - System Menu - Fast Data Entry
      AddAccessControlItem(mvEnv, vFields, "SCFFDE,Fast Data Entry Maintenance,SCSMFD,1000,D,FM")

      'Smart Client - System Menu - DeDuplication
      AddAccessControlItem(mvEnv, vFields, "SCDDCM,Contact Merge,SCSMDD,1000,D,SM")
      AddAccessControlItem(mvEnv, vFields, "SCDDAM,Address Merge,SCSMDD,1100,D,SM")
      AddAccessControlItem(mvEnv, vFields, "SCDDOM,Organisation Merge,SCSMDD,1200,D,SM")
      AddAccessControlItem(mvEnv, vFields, "SCDDAO,Amalgamate Organisations,SCSMDD,1300,D,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDDCD,Contact De-Duplication,SCSMDD,1400,S,SM")
      AddAccessControlItem(mvEnv, vFields, "SCDDBA,Bulk Address Merge,SCSMDD,1500,D,SM")
      AddAccessControlItem(mvEnv, vFields, "SCDDBM,Bulk Contact Merge,SCSMDD,1600,D,SM")
      AddAccessControlItem(mvEnv, vFields, "SCDDBO,Bulk Organisation Merge,SCSMDD,1700,D,SM")
      AddAccessControlItem(mvEnv, vFields, "SCDDPD,Process Duplicates,SCSMDD,1800,S,SM")

      'Smart Client - System Manager - Direct Debits
      AddAccessControlItem(mvEnv, vFields, "SCFDMF,Mandate File Creation,SCSMDB,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDDB,Create Direct Debit Batches,SCSMDB,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDCF,Claim File Creation,SCSMDB,1200,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDCB,Upload BACS Messaging Data,SCSMDB,1400,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDRF,Credit File Creation,SCSMDB,1300,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDBR,Process BACS Messaging,SCSMDB,1400,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDCM,Convert Manual Direct Debits,SCSMDB,1500,S,FM")

      'Smart Client - System Menu - Distribution Boxes
      AddAccessControlItem(mvEnv, vFields, "SCFDBU,Create Unallocated Boxes,SCSMDI,1000,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDST,Print Thank You Letters,SCSMDI,1200,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDSN,Print Advice Notes,SCSMDI,1300,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDSP,Print Packing Slips,SCSMDI,1400,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDSL,Print Box Labels,SCSMDI,1500,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDSB,Set Shipping Information,SCSMDI,1600,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDSA,Set Arrival Information,SCSMDI,1700,N,FM")


      'Smart Client - System Menu - Distribution Boxes Menu - Reports Menu
      AddAccessControlItem(mvEnv, vFields, "SCFDRO,Open Boxes,SCSMDR,1000,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDRU,Unallocated Donations,SCSMDR,1100,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDRA,Allocated Donations,SCSMDR,1200,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDRD,Donor Details,SCSMDR,1300,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDRC,Closed Boxes By Location,SCSMDR,1400,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDRR,Generate Roll of Honour,SCSMDR,1500,N,FM")

      'Smart Client - System Menu - Dutch Electronic Payments
      AddAccessControlItem(mvEnv, vFields, "SCFDPL,Load Payments,SCSMDU,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFDPP,Process Payments,SCSMDU,1100,S,FM")

      'Smart Client - System Menu - Gift Aid Declarations/Covenants
      AddAccessControlItem(mvEnv, vFields, "SCFGDC,Declaration Confirmation,SCSMGA,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFGPC,Create Potential Claim,SCSMGA,1100,U,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFGCT,Create Tax Claim,SCSMGA,1200,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFGCD,Reprint Claim Details Report,SCSMGA,1300,U,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFGCA,Claim Analysis Report,SCSMGA,1400,U,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFGBU,Bulk Update,SCSMGA,1500,S,FM")

      'Smart Client - System Menu - Gift Aid Sponsorship
      AddAccessControlItem(mvEnv, vFields, "SCFGSP,Create Potential Claim,SCSMGS,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFGST,Create Tax Claim,SCSMGS,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFGSC,Reprint Claim Details Report,SCSMGS,1200,S,FM")

      'Smart Client - System Menu - Irish Gift Aid
      AddAccessControlItem(mvEnv, vFields, "SCFGIP,Create Potential Claim,SCSMGI,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFGIT,Create Tax Claim,SCSMGI,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFGIR,Reprint Claim Details Report,SCSMGI,1200,S,FM")

      'Smart Client  - System Menu - Incentives
      AddAccessControlItem(mvEnv, vFields, "SCFIMI,Maintain Incentives,SCSMIN,1000,U,FM")

      'Smart Client - System Menu - Mailings
      AddAccessControlItem(mvEnv, vFields, "SCMMRE,Produce Mailing Documents,SCSMMA,1000,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMFD,Find Mailing Documents,SCSMMA,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMTL,Thank You Letters,SCSMMA,1200,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMFM,Find Mailings,SCSMMA,1300,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMEP,EMail Processor,SCSMMA,1400,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMAC,List All Contacts,SCSMMA,1500,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMDB,Direct Debit,SCSMMA,1600,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMGA,Irish Gift Aid,SCSMMA,1700,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMMB,Members,SCSMMA,1800,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMPA,Payers,SCSMMA,1900,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMPG,PreTax Payroll Giving Pledges,SCSMMA,2000,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMSM,Selection Manager,SCSMMA,2100,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMST,Selection Tester,SCSMMA,2200,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMSO,Standing Orders,SCSMMA,2300,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMSC,Subscriptions,SCSMMA,2400,D,FM")

      'Smart Client - System Menu - Mailings - Events
      AddAccessControlItem(mvEnv, vFields, "SCMMEB,Event Booking,SCSMMA,1100,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMED,Event Delegates,SCSMMA,1200,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMEN,Event Personnel,SCSMMA,1300,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMES,Event Sponsors,SCSMMA,1400,D,FM")

      'Smart Client - System Menu - Mailings - Exams
      AddAccessControlItem(mvEnv, vFields, "SCMMXB,Exam Booking,SCSMMA,1100,N,FM")
      AddAccessControlItem(mvEnv, vFields, "SCMMXC,Exam Candidates,SCSMMA,1200,N,FM")

      'Smart Client - System Menu - Marketing
      AddAccessControlItem(mvEnv, vFields, "SCMKGD,Generate Data,SCSMMK,1000,S,MA")

      'Smart Client - System Menu - Membership
      AddAccessControlItem(mvEnv, vFields, "SCMEFM,Future Membership Changes,SCSMME,1000,D,ME")
      AddAccessControlItem(mvEnv, vFields, "SCMEMC,Membership Cards,SCSMME,1100,U,ME")
      AddAccessControlItem(mvEnv, vFields, "SCMEMS,Membership Suspension,SCSMME,1200,S,ME")
      AddAccessControlItem(mvEnv, vFields, "SCMENM,New Member Fulfilment,SCSMME,1300,D,ME")

      'Smart Client - System Menu - Membership Reports
      AddAccessControlItem(mvEnv, vFields, "SCMRAV,Assumed Voting Rights Report,SCSMRP,1000,S,ME")
      AddAccessControlItem(mvEnv, vFields, "SCMRBP,Ballot Paper Production,SCSMRP,1100,S,ME")
      AddAccessControlItem(mvEnv, vFields, "SCMRBD,Branch Donations Report,SCSMRP,1200,S,ME")
      AddAccessControlItem(mvEnv, vFields, "SCMRBI,Branch Income Report,SCSMRP,1300,S,ME")
      AddAccessControlItem(mvEnv, vFields, "SCMRJA,Junior Analysis Report,SCSMRP,1400,S,ME")

      'Smart Client - System Menu - Membership Statistics
      AddAccessControlItem(mvEnv, vFields, "SCMSGD,Generate Data,SCSMST,1000,S,ME")
      AddAccessControlItem(mvEnv, vFields, "SCMSDR,Detailed Report,SCSMST,1100,S,ME")
      AddAccessControlItem(mvEnv, vFields, "SCMSSR,Summary Report,SCSMST,1200,S,ME")

      'Smart Client - System Menu - Nominal Codes
      AddAccessControlItem(mvEnv, vFields, "SCFNSR,Summary Report,SCSMNC,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFNDR,Detailed Report,SCSMNC,1100,S,FM")

      'Smart Client - System Menu - Paying in slips
      AddAccessControlItem(mvEnv, vFields, "SCFPCL,Load Bank Statement Data,SCSMPI,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPCR,Automated Reconciliation,SCSMPI,1100,S,FM")

      'Smart Client - System Menu - Payment Plans
      AddAccessControlItem(mvEnv, vFields, "SCFPRR,Renewals & Reminders,SCSMPP,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPRO,Remove Old Details Arrears,SCSMPP,1100,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPCE,Cancel Expired Payment Plans,SCSMPP,1200,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPNM,Non-member Fulfilment,SCSMPP,1300,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPCP,Update Products,SCSMPP,1500,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPOS,Apply Surcharges,SCSMPP,1600,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPLI,Re-calculate Loan Interest,SCSMPP,1700,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPUL,Update Loan Interest Rates,SCSMPP,1800,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFEPP,Transfer Payment Plan Changes,SCSMPP,1900,S,FM")

      'Smart Client - System Menu - Payroll Giving
      AddAccessControlItem(mvEnv, vFields, "SCFYPL,Load Pre Tax Payment Data,SCSMPG,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFYRC,Pre Tax Auto Reconciliation,SCSMPG,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFYBC,Pre Tax Bulk Cancellation,SCSMPG,1200,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFYPT,Post Tax Auto Reconciliation,SCSMPG,1300,S,FM")

      'Smart Client - System Menu - Purchase Orders
      AddAccessControlItem(mvEnv, vFields, "SCFPPP,Transfer Payments,SCSMPR,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPTS,Transfer Suppliers,SCSMPR,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPRA,Authorise Payments,SCSMPR,1200,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPPS,Process Payments,SCSMPR,1300,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPAG,Auto Generate,SCSMPR,1400,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPPO,Print,SCSMPR,1500,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPQP,Cheque Production,SCSMPR,1600,S,FM")

      'Smart Client - System Menu - Products
      AddAccessControlItem(mvEnv, vFields, "SCFPPC,Price Change Update,SCSMPD,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFPPR,Purchased Product Report,SCSMPD,1100,S,FM")

      'Smart Client - System Menu - Standing Orders
      AddAccessControlItem(mvEnv, vFields, "SCFSLB,Load Bank Statement Data,SCSMSO,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSAR,Automated Reconciliation,SCSMSO,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSBC,Bulk Cancellation,SCSMSO,1200,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSMR,Manual Reconciliation,SCSMSO,1300,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSRR,Reconciliation Report,SCSMSO,1400,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSBT,Bank Transactions Report,SCSMSO,1500,S,FM")

      'Smart Client - System Menu - Stock
      AddAccessControlItem(mvEnv, vFields, "SCFSPL,Picking List Production,SCSMSK,1000,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSCA,Confirm Stock Allocation,SCSMSK,1100,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSSB,Allocate Stock to Back Orders,SCSMSK,1200,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSDN,Despatch Notes,SCSMSK,1300,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSDT,Despatch Tracking,SCSMSK,1400,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSBO,Back Orders Report,SCSMSK,1500,U,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSSA,Sales Analysis MYL,SCSMSK,1550,U,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSSD,Sales Analysis Detailed,SCSMSK,1570,U,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSSS,Sales Analysis Summary,SCSMSK,1600,U,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSSM,Stock Movement,SCSMSK,1750,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSTP,Transfer Stock To Pack,SCSMSK,1775,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSES,Export Stock,SCSMSK,1800,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSPB,Purge Back Orders,SCSMSK,2000,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSPP,Purge Picking & Despatch Data,SCSMSK,2100,D,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSAC,Picking Lists Awaiting Confirm,SCSMSK,2200,S,FM")
      AddAccessControlItem(mvEnv, vFields, "SCFSSV,Stock Valuation Report,SCSMSK,2300,S,FM")

      'Smart Client - System Menu - Updates
      AddAccessControlItem(mvEnv, vFields, "SCAMAH,Amendment History,SCSMUP,1000,S,SM")
      AddAccessControlItem(mvEnv, vFields, "SCAMPA,Process Address Changes,SCSMUP,1100,S,SM")
      AddAccessControlItem(mvEnv, vFields, "SCAMPC,Set Post Dated Contacts,SCSMUP,1200,S,SM")
      AddAccessControlItem(mvEnv, vFields, "SCAMUA,Update Action Status,SCSMUP,1300,S,SM")
      AddAccessControlItem(mvEnv, vFields, "SCAMUM,Update Mailsort,SCSMUP,1400,D,SM")
      AddAccessControlItem(mvEnv, vFields, "SCAUPU,Update Principal User,SCSMUP,1400,S,SM")
      AddAccessControlItem(mvEnv, vFields, "SCAMRD,Update Regional Data,SCSMUP,1450,S,SM")
      AddAccessControlItem(mvEnv, vFields, "SCAMSN,Update Search Names,SCSMUP,1500,S,SM")
      AddAccessControlItem(mvEnv, vFields, "SCAMPV,Postcode Validation,SCSMUP,1550,S,SM")
      AddAccessControlItem(mvEnv, vFields, "SCAMPS,Purge Sticky Notes,SCSMUP,1600,S,SM")

      'Smart Client - Administration Menu
      AddAccessControlItem(mvEnv, vFields, "SCAMAC,Access Control,SCAM,1000,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMCM,Configuration Maintenance,SCAM,1100,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMLM,Licence Maintenance,SCAM,1200,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMMS,Maintenance Setup,SCAM,1300,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMOM,Ownership Maintenance,SCAM,1400,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMRM,Report Maintenance,SCAM,1500,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMTM,Trader Application Maintenance,SCAM,1600,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMDU,Database Upgrade,SCAM,1700,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMDI,Data Import,SCAM,1800,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMUC,Update Custom Forms,SCAM,1900,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMUG,Update Government Regions,SCAM,2000,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMUD,Update Mailsort Data,SCAM,2100,D,CD")    'aaded new update mailsort data.....
      AddAccessControlItem(mvEnv, vFields, "SCAMUP,Update Payment Schedule,SCAM,2200,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMUT,Update Trader Applications,SCAM,2300,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMPU,Postcode Update,SCAM,2400,N,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMDP,Data Updates,SCAM,2500,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMIT,Import Trader Application,SCAM,2600,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMEC,Export Custom Form,SCAM,2700,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMER,Export Report,SCAM,2800,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMET,Export Trader Application,SCAM,2900,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMCR,Configuration Report,SCAM,3000,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMCS,Check Setup,SCAM,3100,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMCP,Check Payment Plans,SCAM,3200,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMRQ,Regenerate Message Queue,SCAM,3300,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCAMEX,Move External Documents,SCAM,3400,D,CD")

      'Smart Client List Manager
      AddAccessControlItem(mvEnv, vFields, "SCLMML,Mail List,SCLM,1000,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCLMSL,Save List,SCLM,1100,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCLMSD,Save Data,SCLM,1200,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCLMRL,Report List,SCLM,1300,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCLMPL,Print Report,SCLM,1400,U,CD")

      'Smart Client - Campaign PopUp Menu
      AddAccessControlItem(mvEnv, vFields, "SCCPNA,New Appeal,SCCP,1000,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPNS,New Segment,SCCP,1100,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPNC,New Collection,SCCP,1200,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPSC,Segment Criteria,SCCP,1300,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPSS,Segment Steps,SCCP,1400,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPSP,Segment Print,SCCP,1500,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPSA,Sum Appeal,SCCP,1600,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPCA,Count Appeal Or Segment,SCCP,1700,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPMA,Mail Appeal,SCCP,1800,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPCP,Copy Data,SCCP,1900,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPCC,Copy Criteria,SCCP,2000,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPRP,Reports,SCCP,2100,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPAC,Actions,SCCP,2200,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPCI,Calculate Income,SCCP,2300,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPPA,Paste,SCCP,2400,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPAB,Add Collection Boxes,SCCP,2500,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPCO,Count Collectors,SCCP,2600,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPCF,Collection Fulfilment,SCCP,2700,S,MA")
      AddAccessControlItem(mvEnv, vFields, "SCCPMP,Save Mail Appeal Parameters,SCCP,2800,S,MA")

      'Smart Client - Selection Set PopUp Menu
      AddAccessControlItem(mvEnv, vFields, "SCSSDA,Delete All Contacts,SCSS,1000,S,CD")

      'Smart Client - Browser PopUp Menu
      AddAccessControlItem(mvEnv, vFields, "SCBMNE,New,SCBM,1000,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMEE,Edit,SCBM,1100,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMED,Edit Details,SCBM,1200,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMDE,Details,SCBM,1300,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMSN,Sticky Notes,SCBM,1400,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMJO,Journal,SCBM,1500,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMSE,Send EMail,SCBM,1600,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMDN,Dial Number,SCBM,1700,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMSU,Suppressions,SCBM,1800,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMCM,Communications,SCBM,1900,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMAC,Actions,SCBM,2000,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMAT,Activities,SCBM,2100,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMRE,Relationships,SCBM,2200,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMRP,Report,SCBM,2300,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMRS,Reports,SCBM,2400,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMSS,Set Status,SCBM,2500,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMAS,Action Schedule,SCBM,2600,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMCO,Convert To Organisation,SCBM,2700,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMOC,Clone Organisation,SCBM,2800,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMRI,Remove,SCBM,2900,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCBMAF,Add To Favourites,SCBM,3000,U,CD")
      AddAccessControlItem(mvEnv, vFields, "SCSSME,Merge,SCSS,1000,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCSSRN,Rename,SCSS,1000,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCSSCP,Copy,SCSS,1000,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCSSSS,Save As Selection Set,SCSS,1000,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCSSLM,Go To List Manager,SCSS,1000,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCSSDE,Delete,SCSS,1000,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCSSSR,SurveyRegistration,SCSS,1000,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCSSUA,Bulk Update Activity,SCSS,1000,S,CD")

      'Smart Client - Preferences
      AddAccessControlItem(mvEnv, vFields, "SCPRMA,Modify Appearance,SCPR,1000,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCPRMF,Modify Fonts,SCPR,1100,D,CD")

      'Smart Client - Financial PopUp Menu
      AddAccessControlItem(mvEnv, vFields, "CDFPCM,Change Membership Type,CDFP,1000,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPCN,Cancel,CDFP,1100,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPFC,Future Cancellation,CDFP,1200,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPCC,Change Cancellation Reason,CDFP,1300,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPRN,Reprint Numbers,CDFP,1400,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPRC,Flag Memb Card for Reprint,CDFP,1500,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPPC,Payment Plan Conversion,CDFP,1600,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPPM,Payment Plan Maintenance,CDFP,1700,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPPP,Payment Plan Print,CDFP,1800,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPRM,Reinstate Membership,CDFP,1900,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPME,Add Member,CDFP,2000,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPAM,Financial History Change Payer,CDFP,2100,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPAR,Financial Adjustment Reverse,CDFP,2200,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPAF,Financial Adjustment Refund,CDFP,2300,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPAA,Financial Adjustment Analysis,CDFP,2400,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPFM,Future Membership Type,CDFP,2500,U,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPCT,Confirm Transaction,SCFP,2600,S,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPCP,Payment Plan Change Payer,CDFP,2700,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPRD,Reinstate Auto Payment Method,CDFP,2800,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPSP,Skip Payment,CDFP,2900,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPAG,Add Gift Aid Declaration,CDFP,3000,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPAD,Advance Renewal Date,CDFP,3100,S,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPNP,Confirm Payment Plan,CDFP,3200,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPRP,Reinstate Payment Plan,CDFP,3300,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPSC,Change Deliver To,CDFP,3400,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPRL,Replace Member,CDFP,3500,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDEVFL,Add Financial Link,CDEV,3600,U,EV")
      AddAccessControlItem(mvEnv, vFields, "CDAPAP,Amend Due Date,CDAP,4600,S,CD")
      AddAccessControlItem(mvEnv, vFields, "CDAPAL,Remove Allocations,CDAP,4700,S,CD")
      AddAccessControlItem(mvEnv, vFields, "CDFPIF,Payment Plan Refund InAdvance,CDFP,4800,U,FP")
      AddAccessControlItem(mvEnv, vFields, "CDFPIR,Payment Plan Reverse InAdvance,CDFP,4900,U,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPEN,Edit Financial History Notes,SCFP,5000,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPEB,Amend Event Booking,SCFP,5100,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPRC,Reissue Cheque,SCFP,5200,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPUF,Unlock Fundraising Request,SCFP,5300,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPAF,Add Fundraising Payment Link,SCFP,5600,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPEU,Edit Transaction Notes,SCFP,5700,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPCQ,Change Cheque Payee,SCFP,5800,S,FP")
      AddAccessControlItem(mvEnv, vFields, "CDEVCA,Cancel Event Booking,CDEV,5900,U,EV")
      AddAccessControlItem(mvEnv, vFields, "CDEVSI,Supplementary Information,CDEV,6000,U,EV")
      AddAccessControlItem(mvEnv, vFields, "CDEVWL,Waiting List Management,CDEV,6050,U,EV")
      AddAccessControlItem(mvEnv, vFields, "SCFPAP,Authorise Purchase Order,CDFP,6100,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPAR,Add Payment Receipt,CDFP,6100,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPSS,Set Cheque Status,CDFP,6200,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPCI,Re-calculate Interest,SCFP,6300,U,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPCX,Cancel Exam Booking,SCFP,6400,U,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPXC,Change Exam Booking Centre,SCFP,6500,U,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPER,Edit History Reference,SCFP,6600,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPMC,Produce Membership Card,SCFP,6700,U,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPDT,Display Transactions,SCFP,6800,U,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPCA,Change Invoice Address,SCFP,6900,U,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPPI,Preview Invoice,SCFP,7000,U,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPOP,Cancel Purchase Order Payment,SCFP,7100,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPAA,Authorise PO Payment,SCFP,7200,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPPA,Reanalyse PO Payment,SCFP,7300,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPFR,Future Member Renewal Amount,SCFP,7400,U,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPCD,Change Claim Date,SCFP,7500,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCFPGR,Generate Receipt,SCFP,7600,U,FP")

      'Smart Client - Purchase Orders PopUp menu
      AddAccessControlItem(mvEnv, vFields, "SCPOPC,Cancel Purchase Order,SCPO,1000,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCPOPR,Reinstate Purchase Order,SCPO,1100,S,FP")
      AddAccessControlItem(mvEnv, vFields, "SCPOPA,Amend Purchase Order,SCPO,1200,S,FP")

      'Smart Client - Dashboard
      AddAccessControlItem(mvEnv, vFields, "SCDAMA,Dashboard Maintenance,SCDA,1000,U,CD")

      'Smart Client - Dashboard Popup Menu
      AddAccessControlItem(mvEnv, vFields, "SCDPGI,Add Grid,SCDP,1000,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPGR,Add Graph,SCDP,1100,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPGU,Add Guage,SCDP,1200,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPAB,Add Bar,SCDP,1300,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPAD,Add Display Panel,SCDP,1400,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPAW,Add Web Page,SCDP,1500,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPDC,Add Data Chart,SCDP,1600,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPRE,Revert,SCDP,1700,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPDS,Delete System Dashboard,SCDP,1800,D,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPDD,Delete Department Dashboard,SCDP,1900,S,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPDU,Delete User Dashboard,SCDP,2000,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPSY,Save System Dashboard,SCDP,2100,D,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPDE,Save Department Dashboard,SCDP,2200,S,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPUS,Save User Dashboard,SCDP,2300,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPSS,Save As System Dashboard,SCDP,2400,D,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPSD,Save As Department Dashboard,SCDP,2500,S,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPSU,Save As User Dashboard,SCDP,2600,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPHE,Show Header,SCDP,2700,U,SC")
      AddAccessControlItem(mvEnv, vFields, "SCDPFO,Show Footer,SCDP,2800,U,SC")

      'Smart Client - Mailing PopUp Menu
      AddAccessControlItem(mvEnv, vFields, "SCMPDM,Delete Mailing Document,SCMP,1000,S,CD")
      AddAccessControlItem(mvEnv, vFields, "SCMPUM,Unfulfill Mailing Document,SCMP,1100,S,CD")

      'Smart Client - Customise PopUp Menu
      AddAccessControlItem(mvEnv, vFields, "SCCUCU,Customise for Maint. Panels,SCCU,1000,D,CD")
      AddAccessControlItem(mvEnv, vFields, "SCCURV,Revert for Maint. Panels,SCCU,1100,D,CD")

      'Smart Client - Document PopUp Menu
      AddAccessControlItem(mvEnv, vFields, "CDDPND,New Document,CDDP,1000,U,CD")
      AddAccessControlItem(mvEnv, vFields, "CDDPUP,Edit Document,CDDP,1100,U,CD")
      AddAccessControlItem(mvEnv, vFields, "CDDPVD,View Document Content,CDDP,1200,U,CD")
      AddAccessControlItem(mvEnv, vFields, "CDDPED,Edit Document Content,CDDP,1300,U,CD")
      AddAccessControlItem(mvEnv, vFields, "CDDPPR,Print Document Content,CDDP,1400,U,CD")
      AddAccessControlItem(mvEnv, vFields, "CDDPPD,Print Details,CDDP,1500,U,CD")
      AddAccessControlItem(mvEnv, vFields, "CDDPMN,Mark Notified,CDDP,1600,U,CD")
      AddAccessControlItem(mvEnv, vFields, "CDDPMP,Mark Processed,CDDP,1700,U,CD")
      AddAccessControlItem(mvEnv, vFields, "CDDPDE,Delete Document,CDDP,1800,U,CD")
      AddAccessControlItem(mvEnv, vFields, "CDDPLD,Show Related Documents,CDDP,1900,U,CD")
      AddAccessControlItem(mvEnv, vFields, "CDDPAL,New Link to Existing Document,CDDP,2000,U,CD")
      AddAccessControlItem(mvEnv, vFields, "CDDPDL,Delete Document Link,CDDP,2100,U,CD")

      'Smart Client - Event Maintenance
      AddAccessControlItem(mvEnv, vFields, "SCEMUD,Update Department,SCEM,1000,D,EV")
      AddAccessControlItem(mvEnv, vFields, "CDEVSC,Event CPD Maintenance,SCEM,1100,D,EV")

      'Smart Client - Contacts
      AddAccessControlItem(mvEnv, vFields, "CDCMDE,Delete,SCCO,1000,S,CD")
      AddAccessControlItem(mvEnv, vFields, "CDCMCS,Change Status,SCCO,1100,S,CD")
      AddAccessControlItem(mvEnv, vFields, "CDCMCD,Change Department,SCCO,1200,S,CD")
      AddAccessControlItem(mvEnv, vFields, "CDCMCR,Change Source,SCCO,1300,S,CD")
      AddAccessControlItem(mvEnv, vFields, "CDCMCO,Change Ownership Details,SCCO,1400,S,CD")
      AddAccessControlItem(mvEnv, vFields, "CDCMER,Edit Exam Results,SCCO,1500,N,CD")

      'Smart Client - General
      AddAccessControlItem(mvEnv, vFields, "AMAMAN,Table Maintenance Admin Notes,SCGE,1000,S,SM")
      AddAccessControlItem(mvEnv, vFields, "AMAMDL,DisplayList Maint. Grid/Panels,SCGE,1100,S,SM")
      AddAccessControlItem(mvEnv, vFields, "AMFMIC,Import CheckControls CheckBox,SCGE,1200,S,SM")
      AddAccessControlItem(mvEnv, vFields, "GETMST,Schedule Tasks,SCGE,1300,S,GE")
      AddAccessControlItem(mvEnv, vFields, "RESEXP,Restrict Grid Export,SCGE,1400,S,GE")
      AddAccessControlItem(mvEnv, vFields, "GESHTO,Show Total On Grids,SCGE,1400,S,GE")

      'Smart Client - System Menu - Events
      AddAccessControlItem(mvEnv, vFields, "SCEVBB,Block Booking,SCSMEV,1000,U,SM")
      AddAccessControlItem(mvEnv, vFields, "SCEVCN,Cancel,SCSMEV,1100,U,SM")
      AddAccessControlItem(mvEnv, vFields, "SCECPB,Cancel Provisional Bookings,SCSMEV,1200,U,SM")

      'Smart Client - System Menu - Exams
      AddAccessControlItem(mvEnv, vFields, "SCEXMA,Exam Maintenance,SCSMEX,1000,N,SM")
      AddAccessControlItem(mvEnv, vFields, "SCEXER,Enter Exam Results,SCSMEX,1100,N,SM")
      AddAccessControlItem(mvEnv, vFields, "SCEXAC,Allocate Candidate Numbers,SCSMEX,1200,N,SM")
      AddAccessControlItem(mvEnv, vFields, "SCEXAM,Allocate Markers,SCSMEX,1300,N,SM")
      AddAccessControlItem(mvEnv, vFields, "SCEXAG,Apply Grading,SCSMEX,1400,N,SM")
      AddAccessControlItem(mvEnv, vFields, "SCEXGE,Generate Exemption Invoices,SCSMEX,1500,N,SM")
      AddAccessControlItem(mvEnv, vFields, "SCEXLC,Load CSV Results,SCSMEX,1600,N,SM")
      AddAccessControlItem(mvEnv, vFields, "SCEXCP,Cancel Provisional Bookings,SCSMEX,1700,N,SM")
      AddAccessControlItem(mvEnv, vFields, "SCEXPC,Process Certificates,SCSMEX,1800,N,SM")
      AddAccessControlItem(mvEnv, vFields, "SCEXGC,Generate Certificates,SCSMEX,1900,N,SM")

      'Smart Client -System Menu- CPD
      AddAccessControlItem(mvEnv, vFields, "SCCPAP,Apply Points,SCSMCP,1000,U,SM")

      'Smart Client - Mailings
      AddAccessControlItem(mvEnv, vFields, "GMMHCM,Optional Mailing History,SCGM,1000,N,GM")

      'Extra
      AddAccessControlItem(mvEnv, vFields, "CDFPPR,Part Refund,CDFP,1000,U,FP")

      'Exam Maintenance
      'Cources
      AddAccessControlItem(mvEnv, vFields, "SCXMBC,Courses Button,SCXACO,1000,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMCA,Assessment Types,SCXACO,1100,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMCG,Grading,SCXACO,1200,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMCM,Marker Allocation,SCXACO,1300,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMCP,Personnel,SCXACO,1400,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMCQ,Prerequisites,SCXACO,1500,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMCR,Requirements,SCXACO,1600,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMCS,Resources,SCXACO,1700,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMCH,Schedule,SCXACO,1800,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMCC,Categories,SCXACO,1900,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMSM,Study Modes,SCXACO,2000,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMRT,Certificates,SCXACO,2100,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMCD,Documents,SCXACO,2200,S,EX")

      'Centres
      AddAccessControlItem(mvEnv, vFields, "SCXMBE,Centres Button,SCXACE,1000,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMET,Actions,SCXACE,1100,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMEA,Assessment Types,SCXACE,1200,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMEC,Contacts,SCXACE,1300,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMEO,Courses,SCXACE,1400,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMCT,Categories,SCXACE,1800,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMED,Documents,SCXACE,1800,S,EX")

      'Personnel
      AddAccessControlItem(mvEnv, vFields, "SCXMBP,Personnel Button,SCXAPE,1000,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMPA,Assessment Types,SCXAPE,1100,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMPX,Expenses,SCXAPE,1200,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMPM,Marker Information,SCXAPE,1300,S,EX")
      'Sessions
      AddAccessControlItem(mvEnv, vFields, "SCXMBS,Sessions Button,SCXASE,1000,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMSC,Centres,SCXASE,1100,S,EX")
      'Exemptions
      AddAccessControlItem(mvEnv, vFields, "SCXMBX,Exemptions Button,SCXAXE,1000,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMXO,Courses,SCXAXE,1100,S,EX")
      'PopUp Menu
      AddAccessControlItem(mvEnv, vFields, "SCXMUS,Search,SCXAPU,1000,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMUC,Create Programme,SCXAPU,1100,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMUR,Reports,SCXAPU,1200,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMMR,Reallocate Marker,SCXAPU,1300,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMMU,Unallocate Marker,SCXAPU,1400,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMSU,Share,SCXAPU,1500,S,EX")
      AddAccessControlItem(mvEnv, vFields, "SCXMSD,Clone,SCXAPU,1600,S,EX")

    End Sub

    Private Sub AddAccessControlArea(ByVal pEnv As CDBEnvironment, ByVal pFields As CDBFields, ByVal pValues As String)
      Dim vValues() As String
      vValues = pValues.Split(","c) 'Split(pValues, ",")
      pFields(3).Value = vValues(0)
      pFields(4).Value = vValues(1)
      pFields(5).Value = vValues(2)
      pEnv.Connection.InsertRecord("access_control_areas", pFields)
    End Sub

    Private Sub AddAccessControlItem(ByVal pEnv As CDBEnvironment, ByVal pFields As CDBFields, ByVal pValues As String)
      Dim vValues() As String
      Dim vIndex As Integer
      vValues = pValues.Split(","c)
      For vIndex = 0 To 5
        pFields(vIndex + 3).Value = vValues(vIndex)
      Next
      pEnv.Connection.InsertRecord("access_control_items", pFields)
    End Sub

  End Class

End Namespace
