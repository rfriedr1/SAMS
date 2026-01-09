object dm: Tdm
  Height = 866
  Width = 1228
  object qrySubmit: TADOQuery
    Parameters = <>
    Left = 632
    Top = 632
  end
  object qryDB: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    Left = 56
    Top = 80
  end
  object qrySampleOfSubmitter: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 56
    Top = 192
  end
  object dsSampleOfSubmitter: TDataSource
    DataSet = qrySampleOfSubmitter
    Left = 120
    Top = 176
  end
  object qryAllUser: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'SELECT user_nr, first_name, last_name, organisation from user_t ' +
        'where correspondance=1 ORDER BY last_name;')
    Left = 728
    Top = 632
  end
  object dsQryDb: TDataSource
    DataSet = qryDB
    Left = 120
    Top = 80
  end
  object qryProjectsOfUser: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    Left = 56
    Top = 260
  end
  object dsProjectsOfUser: TDataSource
    DataSet = qryProjectsOfUser
    Left = 136
    Top = 260
  end
  object qrySamplesOfProject: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 56
    Top = 316
  end
  object dsSamplesOfProject: TDataSource
    DataSet = qrySamplesOfProject
    Left = 136
    Top = 300
  end
  object qrySample: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 56
    Top = 364
  end
  object qryTarget: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 56
    Top = 416
  end
  object dsTarget: TDataSource
    DataSet = qryTarget
    Left = 144
    Top = 416
  end
  object dsSample: TDataSource
    DataSet = qrySample
    Left = 144
    Top = 364
  end
  object adoConnKTL: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=MSDASQL.1;Persist Security Info=False;Data Source=db_dm' +
      'ams;'
    ConnectionTimeout = 25
    DefaultDatabase = 'db_dmams'
    LoginPrompt = False
    OnExecuteComplete = adoConnKTLExecuteComplete
    Left = 55
    Top = 16
  end
  object tblUser: TADOTable
    Connection = adoConnKTL
    CursorType = ctStatic
    TableName = 'user_t'
    Left = 389
    Top = 126
    object tblUsersalutation: TStringField
      FieldName = 'salutation'
    end
    object tblUseruser_nr: TAutoIncField
      FieldName = 'user_nr'
      ReadOnly = True
    end
    object tblUserfirst_name: TStringField
      DisplayWidth = 15
      FieldName = 'first_name'
      Size = 40
    end
    object tblUserlast_name: TStringField
      DisplayWidth = 20
      FieldName = 'last_name'
      Size = 60
    end
    object tblUserorganisation: TStringField
      FieldName = 'organisation'
      Size = 99
    end
    object tblUserinstitute: TStringField
      FieldName = 'institute'
      Size = 99
    end
    object tblUseraddress_1: TStringField
      FieldName = 'address_1'
      Size = 40
    end
    object tblUseraddress_2: TStringField
      FieldName = 'address_2'
      Size = 40
    end
    object tblUsertown: TStringField
      FieldName = 'town'
      Size = 30
    end
    object tblUserpostcode: TStringField
      FieldName = 'postcode'
      Size = 10
    end
    object tblUsercountry: TStringField
      FieldName = 'country'
      Size = 30
    end
    object tblUserphone_1: TStringField
      FieldName = 'phone_1'
      Size = 30
    end
    object tblUserphone_2: TStringField
      FieldName = 'phone_2'
      Size = 30
    end
    object tblUserfax: TStringField
      FieldName = 'fax'
      Size = 30
    end
    object tblUseremail: TStringField
      FieldName = 'email'
      Size = 80
    end
    object tblUserwww: TStringField
      FieldName = 'www'
      Size = 60
    end
    object tblUseraccount: TStringField
      FieldName = 'account'
      Size = 40
    end
    object tblUserinvoice: TSmallintField
      FieldName = 'invoice'
    end
    object tblUsercorrespondance: TSmallintField
      FieldName = 'correspondance'
    end
    object tblUseruser_comment: TMemoField
      FieldName = 'user_comment'
      BlobType = ftMemo
    end
    object tblUserlanguage: TStringField
      FieldName = 'language'
      Size = 2
    end
  end
  object tblProjects: TADOTable
    Connection = adoConnKTL
    CursorType = ctStatic
    TableName = 'project_t'
    Left = 389
    Top = 206
  end
  object dsProjects: TDataSource
    DataSet = tblProjects
    Left = 469
    Top = 206
  end
  object tblTypes: TADOTable
    Connection = adoConnKTL
    CursorType = ctStatic
    TableName = 'sampletype_t'
    Left = 389
    Top = 254
  end
  object dsTypes: TDataSource
    DataSet = tblTypes
    Left = 469
    Top = 254
  end
  object tblMaterial: TADOTable
    Connection = adoConnKTL
    CursorType = ctStatic
    TableName = 'material_t'
    Left = 391
    Top = 302
  end
  object dsMaterial: TDataSource
    DataSet = tblMaterial
    Left = 469
    Top = 302
  end
  object tblProjectTypes: TADOTable
    Connection = adoConnKTL
    CursorType = ctStatic
    IndexFieldNames = 'indexnr'
    TableName = 'projecttype_t'
    Left = 389
    Top = 395
  end
  object dsProjectTypes: TDataSource
    DataSet = tblProjectTypes
    Left = 469
    Top = 395
  end
  object tblMethod: TADOTable
    Connection = adoConnKTL
    CursorType = ctStatic
    TableName = 'method_t'
    Left = 389
    Top = 443
    object tblMethodmethod: TStringField
      FieldName = 'method'
    end
    object tblMethoddescr: TMemoField
      FieldName = 'descr'
      OnGetText = tblMethoddescrGetText
      BlobType = ftMemo
    end
    object tblMethodindexnr: TIntegerField
      FieldName = 'indexnr'
    end
  end
  object dsMethod: TDataSource
    DataSet = tblMethod
    Left = 469
    Top = 443
  end
  object tblResearchType: TADOTable
    Connection = adoConnKTL
    CursorType = ctStatic
    IndexFieldNames = 'indexnr'
    TableName = 'research_t'
    Left = 389
    Top = 491
  end
  object dsResearchType: TDataSource
    DataSet = tblResearchType
    Left = 469
    Top = 491
  end
  object tblReportType: TADOTable
    Connection = adoConnKTL
    CursorType = ctStatic
    IndexFieldNames = 'indexnr'
    TableName = 'reporttype_t'
    Left = 389
    Top = 539
  end
  object dsReportType: TDataSource
    DataSet = tblReportType
    Left = 469
    Top = 539
  end
  object qrySamplesAvailable: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 48
    Top = 480
  end
  object dsSamplesAvailable: TDataSource
    DataSet = qrySamplesAvailable
    Left = 144
    Top = 480
  end
  object adoCmd: TADOCommand
    Connection = adoConnKTL
    Parameters = <>
    Left = 388
    Top = 56
  end
  object qryPrepSteps: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 219
    Top = 81
  end
  object dsPrepSteps: TDataSource
    DataSet = qryPrepSteps
    Left = 275
    Top = 81
  end
  object qryTargetsAvailable: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 219
    Top = 131
  end
  object dsTargetsAvailable: TDataSource
    DataSet = qryTargetsAvailable
    Left = 275
    Top = 131
  end
  object tblFraction: TADOTable
    Connection = adoConnKTL
    CursorType = ctStatic
    TableName = 'fraction_t'
    Left = 391
    Top = 350
  end
  object dsFraction: TDataSource
    DataSet = tblFraction
    Left = 469
    Top = 350
  end
  object qryTest: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'SELECT sample_t.sample_nr, type, material, fraction, pre_sub_tre' +
        'at, weight, preparation, sampling_date, editable, not_tobedated,' +
        ' user_label, sample_t.user_label_nr, user_desc1, user_desc2, res' +
        'idue,sample_t.user_comment,project_t.project, invoice_nr, in_dat' +
        'e, desired_date, priority, status, price,   user_t.last_name, pr' +
        'eparation_t.prep_nr  FROM sample_t INNER JOIN project_t ON proje' +
        'ct_t.project_nr=sample_t.project_nr INNER JOIN user_t ON user_t.' +
        'user_nr=project_t.user_nr INNER JOIN preparation_t ON preparatio' +
        'n_t.sample_nr=sample_t.sample_nr WHERE sample_t.sample_nr=823 AN' +
        'D preparation_t.prep_nr=1 ;')
    Left = 633
    Top = 256
  end
  object dsTest: TDataSource
    DataSet = qryTest
    Left = 729
    Top = 256
  end
  object qrySampleInfo: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        ' SELECT sample_t.sample_nr, type, material, fraction, pre_sub_tr' +
        'eat, weight, preparation,'
      
        ' sampling_date, editable, not_tobedated, user_label, sample_t.us' +
        'er_label_nr, user_desc1,'
      
        ' user_desc2, residue, sample_t.user_comment,project_t.project, i' +
        'nvoice_nr, in_date,'
      
        ' desired_date, priority, status, price,   user_t.last_name, prep' +
        'aration_t.prep_nr, batch,'
      
        ' prep_comment, weight_start, weight_medium, weight_end, prep_end' +
        ', step1_methode,'
      
        ' step2_methode,step3_methode,step4_methode,step5_methode , targe' +
        't_t.target_nr, cn_ratio,'
      
        ' c_percent, n_percent, preparation_t.stop, magazine, position, p' +
        'recis, cycle_min, cycle_max,'
      
        ' catalyst, cathode_nr, reactor_nr, co2_init, co2_final, hydro_in' +
        'it, hydro_final, react_time,'
      
        ' target_pressed, target_t.stop, fm, fm_sig, dc13, dc13_sig,calcs' +
        'et, editallowed,'
      
        ' target_t.c14_age, target_t.c14_age_sig, target_t.cal1sMin, targ' +
        'et_t.cal1sMax, target_t.cal2sMin,'
      
        ' target_t.cal2sMax  FROM sample_t INNER JOIN project_t ON projec' +
        't_t.project_nr=sample_t.project_nr'
      ' INNER JOIN user_t ON user_t.user_nr=project_t.user_nr'
      
        ' INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sa' +
        'mple_nr'
      ' INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr'
      
        ' WHERE sample_t.sample_nr=1009 AND preparation_t.prep_nr=1  AND ' +
        'target_nr=1;')
    Left = 633
    Top = 304
  end
  object dsSampleInfo: TDataSource
    DataSet = qrySampleInfo
    Left = 729
    Top = 304
  end
  object tblProjectStatus: TADOTable
    Connection = adoConnKTL
    CursorType = ctStatic
    IndexFieldNames = 'indexnr'
    TableName = 'projectstatus_t'
    Left = 389
    Top = 584
  end
  object dsProjectStatus: TDataSource
    DataSet = tblProjectStatus
    Left = 477
    Top = 592
  end
  object qryPlanned: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        ' SELECT sample_t.sample_nr, type, material, fraction, pre_sub_tr' +
        'eat, weight, preparation,'
      
        ' sampling_date, editable, not_tobedated, user_label, sample_t.us' +
        'er_label_nr, user_desc1,'
      
        ' user_desc2, residue, sample_t.user_comment,project_t.project, i' +
        'nvoice_nr, in_date,'
      
        ' desired_date, priority, status, price,   user_t.last_name, prep' +
        'aration_t.prep_nr, batch,'
      
        ' prep_comment, weight_start, weight_medium, weight_end, prep_end' +
        ', step1_methode,'
      
        ' step2_methode,step3_methode,step4_methode,step5_methode , targe' +
        't_t.target_nr, cn_ratio,'
      
        ' c_percent, n_percent, preparation_t.stop, magazine, position, p' +
        'recis, cycle_min, cycle_max,'
      
        ' catalyst, cathode_nr, reactor_nr, co2_init, co2_final, hydro_in' +
        'it, hydro_final, react_time,'
      
        ' target_pressed, target_t.stop, fm, fm_sig, dc13, dc13_sig,calcs' +
        'et, editallowed,'
      
        ' target_t.c14_age, target_t.c14_age_sig, target_t.cal1sMin, targ' +
        'et_t.cal1sMax, target_t.cal2sMin,'
      
        ' target_t.cal2sMax  FROM sample_t INNER JOIN project_t ON projec' +
        't_t.project_nr=sample_t.project_nr'
      ' INNER JOIN user_t ON user_t.user_nr=project_t.user_nr'
      
        ' INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sa' +
        'mple_nr'
      ' INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr'
      
        ' WHERE sample_t.sample_nr=1009 AND preparation_t.prep_nr=1  AND ' +
        'target_nr=1;')
    Left = 633
    Top = 356
  end
  object dsPlanned: TDataSource
    DataSet = qryPlanned
    Left = 729
    Top = 356
  end
  object qryInPrep: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        ' SELECT sample_t.sample_nr, type, material, fraction, pre_sub_tr' +
        'eat, weight, preparation,'
      
        ' sampling_date, editable, not_tobedated, user_label, sample_t.us' +
        'er_label_nr, user_desc1,'
      
        ' user_desc2, residue, sample_t.user_comment,project_t.project, i' +
        'nvoice_nr, in_date,'
      
        ' desired_date, priority, status, price,   user_t.last_name, prep' +
        'aration_t.prep_nr, batch,'
      
        ' prep_comment, weight_start, weight_medium, weight_end, prep_end' +
        ', step1_methode,'
      
        ' step2_methode,step3_methode,step4_methode,step5_methode , targe' +
        't_t.target_nr, cn_ratio,'
      
        ' c_percent, n_percent, preparation_t.stop, magazine, position, p' +
        'recis, cycle_min, cycle_max,'
      
        ' catalyst, cathode_nr, reactor_nr, co2_init, co2_final, hydro_in' +
        'it, hydro_final, react_time,'
      
        ' target_pressed, target_t.stop, fm, fm_sig, dc13, dc13_sig,calcs' +
        'et, editallowed,'
      
        ' target_t.c14_age, target_t.c14_age_sig, target_t.cal1sMin, targ' +
        'et_t.cal1sMax, target_t.cal2sMin,'
      
        ' target_t.cal2sMax  FROM sample_t INNER JOIN project_t ON projec' +
        't_t.project_nr=sample_t.project_nr'
      ' INNER JOIN user_t ON user_t.user_nr=project_t.user_nr'
      
        ' INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sa' +
        'mple_nr'
      ' INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr'
      
        ' WHERE sample_t.sample_nr=1009 AND preparation_t.prep_nr=1  AND ' +
        'target_nr=1;')
    Left = 633
    Top = 422
  end
  object dsInPrep: TDataSource
    DataSet = qryInPrep
    Left = 729
    Top = 422
  end
  object qrySamplesOfLabTask: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    Left = 472
    Top = 80
  end
  object dsSamplesOfLabtask: TDataSource
    DataSet = qrySamplesOfLabTask
    Left = 624
    Top = 8
  end
  object qryDb1: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    Left = 56
    Top = 128
  end
  object dsQryDb1: TDataSource
    DataSet = qryDb1
    Left = 120
    Top = 128
  end
  object qryActiveBatches: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 219
    Top = 202
  end
  object dsActiveBatches: TDataSource
    DataSet = qryActiveBatches
    Left = 283
    Top = 202
  end
  object qryProject: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        ' SELECT sample_t.sample_nr, type, material, fraction, pre_sub_tr' +
        'eat, weight, preparation,'
      
        ' sampling_date, editable, not_tobedated, user_label, sample_t.us' +
        'er_label_nr, user_desc1,'
      
        ' user_desc2, residue, sample_t.user_comment,project_t.project, i' +
        'nvoice_nr, in_date,'
      
        ' desired_date, priority, status, price,   user_t.last_name, prep' +
        'aration_t.prep_nr, batch,'
      
        ' prep_comment, weight_start, weight_medium, weight_end, prep_end' +
        ', step1_methode,'
      
        ' step2_methode,step3_methode,step4_methode,step5_methode , targe' +
        't_t.target_nr, cn_ratio,'
      
        ' c_percent, n_percent, preparation_t.stop, magazine, position, p' +
        'recis, cycle_min, cycle_max,'
      
        ' catalyst, cathode_nr, reactor_nr, co2_init, co2_final, hydro_in' +
        'it, hydro_final, react_time,'
      
        ' target_pressed, target_t.stop, fm, fm_sig, dc13, dc13_sig,calcs' +
        'et, editallowed,'
      
        ' target_t.c14_age, target_t.c14_age_sig, target_t.cal1sMin, targ' +
        'et_t.cal1sMax, target_t.cal2sMin,'
      
        ' target_t.cal2sMax  FROM sample_t INNER JOIN project_t ON projec' +
        't_t.project_nr=sample_t.project_nr'
      ' INNER JOIN user_t ON user_t.user_nr=project_t.user_nr'
      
        ' INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sa' +
        'mple_nr'
      ' INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr'
      
        ' WHERE sample_t.sample_nr=1009 AND preparation_t.prep_nr=1  AND ' +
        'target_nr=1;')
    Left = 631
    Top = 486
  end
  object dsProject: TDataSource
    DataSet = qryProject
    Left = 727
    Top = 486
  end
  object qryProjectsSinceYear: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 56
    Top = 532
  end
  object dsProjectsSinceYear: TDataSource
    DataSet = qryProjectsSinceYear
    Left = 144
    Top = 532
  end
  object qryWeights: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 56
    Top = 583
  end
  object dsWeights: TDataSource
    DataSet = qryWeights
    Left = 144
    Top = 583
  end
  object qryGraphWeight: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 56
    Top = 639
  end
  object dsGraphWeight: TDataSource
    DataSet = qryGraphWeight
    Left = 144
    Top = 631
  end
  object dsUserInfo: TDataSource
    DataSet = tblUser
    Left = 469
    Top = 126
  end
  object dsMagazines: TDataSource
    DataSet = qryMagazines
    Left = 1040
    Top = 528
  end
  object qryMagazines: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    Left = 888
    Top = 528
  end
  object qryMagazineData: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    Left = 888
    Top = 590
  end
  object dsMagazineData: TDataSource
    DataSet = qryMagazineData
    Left = 1040
    Top = 590
  end
  object qryWaitingForGraph: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        ' SELECT sample_t.sample_nr, type, material, fraction, pre_sub_tr' +
        'eat, weight, preparation,'
      
        ' sampling_date, editable, not_tobedated, user_label, sample_t.us' +
        'er_label_nr, user_desc1,'
      
        ' user_desc2, residue, sample_t.user_comment,project_t.project, i' +
        'nvoice_nr, in_date,'
      
        ' desired_date, priority, status, price,   user_t.last_name, prep' +
        'aration_t.prep_nr, batch,'
      
        ' prep_comment, weight_start, weight_medium, weight_end, prep_end' +
        ', step1_methode,'
      
        ' step2_methode,step3_methode,step4_methode,step5_methode , targe' +
        't_t.target_nr, cn_ratio,'
      
        ' c_percent, n_percent, preparation_t.stop, magazine, position, p' +
        'recis, cycle_min, cycle_max,'
      
        ' catalyst, cathode_nr, reactor_nr, co2_init, co2_final, hydro_in' +
        'it, hydro_final, react_time,'
      
        ' target_pressed, target_t.stop, fm, fm_sig, dc13, dc13_sig,calcs' +
        'et, editallowed,'
      
        ' target_t.c14_age, target_t.c14_age_sig, target_t.cal1sMin, targ' +
        'et_t.cal1sMax, target_t.cal2sMin,'
      
        ' target_t.cal2sMax  FROM sample_t INNER JOIN project_t ON projec' +
        't_t.project_nr=sample_t.project_nr'
      ' INNER JOIN user_t ON user_t.user_nr=project_t.user_nr'
      
        ' INNER JOIN preparation_t ON preparation_t.sample_nr=sample_t.sa' +
        'mple_nr'
      ' INNER JOIN target_t ON target_t.sample_nr=sample_t.sample_nr'
      
        ' WHERE sample_t.sample_nr=1009 AND preparation_t.prep_nr=1  AND ' +
        'target_nr=1;')
    Left = 632
    Top = 541
  end
  object dsWaitingForGraph: TDataSource
    DataSet = qryWaitingForGraph
    Left = 728
    Top = 541
  end
  object tblInvoice: TADOTable
    Connection = adoConnKTL
    CursorType = ctStatic
    Filter = 'correspondance=0'
    Filtered = True
    IndexFieldNames = 'last_name'
    TableName = 'user_t'
    Left = 389
    Top = 174
    object AutoIncField1: TAutoIncField
      FieldName = 'user_nr'
      ReadOnly = True
    end
    object StringField1: TStringField
      DisplayWidth = 15
      FieldName = 'first_name'
      Size = 40
    end
    object StringField2: TStringField
      DisplayWidth = 20
      FieldName = 'last_name'
      Size = 60
    end
    object StringField3: TStringField
      FieldName = 'organisation'
      Size = 40
    end
    object StringField4: TStringField
      FieldName = 'institute'
      Size = 40
    end
    object StringField5: TStringField
      FieldName = 'address_1'
      Size = 40
    end
    object StringField6: TStringField
      FieldName = 'address_2'
      Size = 40
    end
    object StringField7: TStringField
      FieldName = 'town'
      Size = 30
    end
    object StringField8: TStringField
      FieldName = 'postcode'
      Size = 10
    end
    object StringField9: TStringField
      FieldName = 'country'
      Size = 30
    end
    object StringField10: TStringField
      FieldName = 'phone_1'
      Size = 30
    end
    object StringField11: TStringField
      FieldName = 'phone_2'
      Size = 30
    end
    object StringField12: TStringField
      FieldName = 'fax'
      Size = 30
    end
    object StringField13: TStringField
      FieldName = 'email'
      Size = 80
    end
    object StringField14: TStringField
      FieldName = 'www'
      Size = 60
    end
    object StringField15: TStringField
      FieldName = 'account'
      Size = 40
    end
    object SmallintField1: TSmallintField
      FieldName = 'invoice'
    end
    object SmallintField2: TSmallintField
      FieldName = 'correspondance'
    end
    object MemoField1: TMemoField
      FieldName = 'user_comment'
      BlobType = ftMemo
    end
  end
  object dsInvoice: TDataSource
    DataSet = tblInvoice
    Left = 469
    Top = 174
  end
  object adoConnectionCEZA: TADOConnection
    ConnectionString = 
      'Provider=MSDASQL.1;Password=Pe02_Du01?;Persist Security Info=Tru' +
      'e;User ID=C14_MAMS;Data Source=db_ma_proben;Initial Catalog=db_m' +
      'a_proben'
    LoginPrompt = False
    Provider = 'MSDASQL.1'
    Left = 960
    Top = 8
  end
  object adoCmdCEZA: TADOCommand
    Connection = adoConnectionCEZA
    Parameters = <>
    Left = 964
    Top = 128
  end
  object qryCEZA: TADOQuery
    Connection = adoConnectionCEZA
    CursorType = ctStatic
    Parameters = <>
    Left = 960
    Top = 64
  end
  object qryProjectsOfReport: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    Left = 216
    Top = 268
  end
  object dsProjectsOfReport: TDataSource
    DataSet = qryProjectsOfReport
    Left = 280
    Top = 268
  end
  object dsCEZA: TDataSource
    DataSet = qryCEZA
    Left = 1025
    Top = 60
  end
  object qryTargetData: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    Left = 56
    Top = 688
  end
  object dsTargetData: TDataSource
    DataSet = qryTargetData
    Left = 144
    Top = 688
  end
  object qryPendingReports: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 632
    Top = 688
  end
  object dsPendingReports: TDataSource
    DataSet = qryPendingReports
    Left = 728
    Top = 685
  end
  object qrySample1: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    Left = 384
    Top = 648
  end
  object qrySample2: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    Left = 480
    Top = 648
  end
  object ADOConnection1: TADOConnection
    Left = 744
    Top = 8
  end
  object qryProjectHasResults: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 632
    Top = 744
  end
  object qryDBPlot: TADOQuery
    Connection = adoConnKTL
    CursorType = ctStatic
    Parameters = <>
    Left = 888
    Top = 272
  end
  object dsDBPlot: TDataSource
    DataSet = qryDBPlot
    Left = 968
    Top = 272
  end
  object qryWaitingForMeas: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 632
    Top = 592
  end
  object dsWaitingForMeas: TDataSource
    DataSet = qryWaitingForMeas
    Left = 728
    Top = 592
  end
  object qryWaitingExpress: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 632
    Top = 800
  end
  object dsWaitingExpress: TDataSource
    DataSet = qryWaitingExpress
    Left = 728
    Top = 800
  end
  object dsLeHeCurrents: TDataSource
    DataSet = cdsLeHeCurrents
    Left = 1008
    Top = 376
  end
  object cdsLeHeCurrents: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 888
    Top = 376
    object cdsLeHeCurrentsSample_nr: TStringField
      FieldName = 'Sample_nr'
    end
    object cdsLeHeCurrentsPrep_nr: TStringField
      FieldName = 'Prep_nr'
    end
    object cdsLeHeCurrentsTarget_nr: TStringField
      FieldName = 'Target_nr'
    end
    object cdsLeHeCurrentsLECurrent: TStringField
      FieldName = 'LECurrent'
    end
    object cdsLeHeCurrentsHECurrent: TStringField
      FieldName = 'HECurrent'
    end
  end
  object qryMagazinesUnpressed: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 888
    Top = 680
  end
  object dsMagazinesUnpressed: TDataSource
    DataSet = qryMagazinesUnpressed
    Left = 1040
    Top = 680
  end
  object qrySamplesOfUnpressedMagazine: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 888
    Top = 744
  end
  object dsSamplesOfUnpressedMagazine: TDataSource
    DataSet = qrySamplesOfUnpressedMagazine
    Left = 1072
    Top = 744
  end
  object qryWaitingForMeasAll: TADOQuery
    Connection = adoConnKTL
    Parameters = <>
    Left = 632
    Top = 168
  end
  object dsWaitingForMeasAll: TDataSource
    DataSet = qryWaitingForMeasAll
    Left = 768
    Top = 168
  end
end
