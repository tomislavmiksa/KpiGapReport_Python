# USED FOR DB ACCESS
import pyodbc
import sqlalchemy
# PYTHON DATA ANALYSIS LIBRARY
import pandas as pd
# WRITING TO EXCEL FILE
# local scripts and libs
import sys
import define_classes as kc
import script_calc_kpi as ck

# CLASES
########################################################################################################################
class OperatorCDRs:
    def __init__(self,server,voice_db,data_db,voice_table,speech_table,data_table, output_table):
        self.srv = server
        self.v_db = voice_db
        self.d_db = data_db
        self.v_cdr = voice_table
        self.s_cdr = speech_table
        self.d_cdr = data_table
        self.table_name = output_table

# MAIN PROGRAM
########################################################################################################################
cdr_views = []

# OPERATOR 1
op = kc.OperatorCDRs(server       = "blndb11",
                     voice_db     = "DE_BM_Voice_1905",
                     data_db      = "DE_BM_Data_1905",
                     voice_table  = "vVoiceCDR2018_Operator1",
                     speech_table = "vSpeechCDR2018_Operator1",
                     data_table   = "vDataCDR2018_Operator1",
                     output_table = 'NEW_KPI_OPERATOR_1')
cdr_views.append(op)
del op
# OPERATOR 2
op = kc.OperatorCDRs(server       = "blndb11",
                     voice_db     = "DE_BM_Voice_1905",
                     data_db      = "DE_BM_Data_1905",
                     voice_table  = "vVoiceCDR2018_Operator2",
                     speech_table = "vSpeechCDR2018_Operator2",
                     data_table   = "vDataCDR2018_Operator2",
                     output_table = 'NEW_KPI_OPERATOR_2')
cdr_views.append(op)
del op
# OPERATOR 2
op = kc.OperatorCDRs(server       = "blndb11",
                     voice_db     = "DE_BM_Voice_1905",
                     data_db      = "DE_BM_Data_1905",
                     voice_table  = "vVoiceCDR2018_Operator3",
                     speech_table = "vSpeechCDR2018_Operator3",
                     data_table   = "vDataCDR2018_Operator3",
                     output_table = 'NEW_KPI_OPERATOR_3')
cdr_views.append(op)
del op

# go through all operators and extract all possible GeoLocation Information
for op in cdr_views:
    print('CONNECTING...\n',
          '\nSERVER            : ',op.srv,
          '\nVOICE DB          : ',op.v_db,
          '\nDATA DB           : ',op.d_db,
          '\nVOICE CDR TABLE   : ',op.v_cdr,
          '\nSPEECH CDR TABLE  : ',op.s_cdr,
          '\nDATA CDR TABLE    : ',op.d_cdr)
    try:
        # DATABASE CONNECTION
        vdb = pyodbc.connect(r'Driver={SQL Server};Server=%s;Database=%s;Trusted_Connection=yes;' %( op.srv, op.v_db ) )
        ddb = pyodbc.connect(r'Driver={SQL Server};Server=%s;Database=%s;Trusted_Connection=yes;' %( op.srv, op.d_db ) )
        vcur = vdb.cursor()
        dcur = ddb.cursor()

        print("DB CONNECTED, STARTING GEOLOCATION EXTRACTION...")

        # GET CDRs TO CALCULATE
        cdr_voice  = pd.read_sql("SELECT DISTINCT G_Level_1,G_Level_2,G_Level_3,G_Level_4,G_Level_5 "
                                 ",CASE WHEN Region       is null THEN '' ELSE Region       END AS Region "
                                 ",CASE WHEN Vendor       is null THEN '' ELSE Vendor       END AS Vendor "
                                 ",CASE WHEN Fleet        is null THEN '' ELSE Fleet        END AS Train "
                                 ",CASE WHEN WagonNumber  is null THEN '' ELSE WagonNumber  END AS Wagon "
                                 "FROM " + op.v_cdr, vdb)
        cdr_voice['G_Level_2_Vendor'] = cdr_voice['G_Level_2'] + ';' + cdr_voice['Vendor']
        cdr_voice['G_Level_3_Vendor'] = cdr_voice['G_Level_3'] + ';' + cdr_voice['Vendor']
        cdr_voice['G_Level_34'] = cdr_voice['G_Level_3'] + ';' + cdr_voice['G_Level_4']
        cdr_voice['G_Level_45'] = cdr_voice['G_Level_4'] + ';' + cdr_voice['G_Level_5']
        cdr_voice['TrainRoute'] = cdr_voice['G_Level_4'] + ';' + cdr_voice['Train'] + ';' + cdr_voice['Wagon']

        # GET CDRs TO CALCULATE
        cdr_data  = pd.read_sql("SELECT DISTINCT G_Level_1,G_Level_2,G_Level_3,G_Level_4,G_Level_5 "
                                 ",CASE WHEN Region       is null THEN '' ELSE Region       END AS Region "
                                 ",CASE WHEN Vendor       is null THEN '' ELSE Vendor       END AS Vendor "
                                 ",CASE WHEN Train_Name   is null THEN '' ELSE Train_Name   END AS Train "
                                 ",CASE WHEN Wagon_Number is null THEN '' ELSE Wagon_Number END AS Wagon "
                                 "FROM " + op.d_cdr, ddb)
        cdr_data['G_Level_2_Vendor'] = cdr_data['G_Level_2'] + ';' + cdr_data['Vendor']
        cdr_data['G_Level_3_Vendor'] = cdr_data['G_Level_3'] + ';' + cdr_data['Vendor']
        cdr_data['G_Level_34'] = cdr_data['G_Level_3'] + ';' + cdr_data['G_Level_4']
        cdr_data['G_Level_45'] = cdr_data['G_Level_4'] + ';' + cdr_data['G_Level_5']
        cdr_data['TrainRoute'] = cdr_data['G_Level_4'] + ';' + cdr_data['Train'] + ';' + cdr_data['Wagon']

        print("\nGEOLOCATION EXTRACTION FROM DB COMPLETED...")

        dc = kc.GeoStuff("Drive", "City")
        dc.updateLocations(cdr_voice)
        dc.updateLocations(cdr_data)

        dr = kc.GeoStuff("Drive", "Connecting Roads")
        dr.updateLocations(cdr_voice)
        dr.updateLocations(cdr_data)

        wc = kc.GeoStuff("Walk", "City")
        wc.updateLocations(cdr_voice)#
        wc.updateLocations(cdr_data)

        wt = kc.GeoStuff("Walk", "Train Route")
        wt.updateLocations(cdr_voice)
        wt.updateLocations(cdr_data)

        # CLOSE DB CONNECTIONS
        vdb.close()
        ddb.close()
        del cdr_voice
        del cdr_data
        print('\n...GEOLOCATION INFORMATION EXTRACTION SUCCESS!')

    except:
        print('\n...GEOLOCATION INFORMATION EXTRACTION FAILED!')
        exit()

print("FOLLOWING GEOLOCATION MODULES EXTRACTED...")
dc.printLocations()
dr.printLocations()
wc.printLocations()
wt.printLocations()

# go through all operators and extract KPI Report
for op in cdr_views:
    print('\nKPI CALCULATION FOR OPERATOR...\n')
    print('\nCONNECTING...\n',
          '\nSERVER            : ',op.srv,
          '\nVOICE DB          : ',op.v_db,
          '\nDATA DB           : ',op.d_db,
          '\nVOICE CDR TABLE   : ',op.v_cdr,
          '\nSPEECH CDR TABLE  : ',op.s_cdr,
          '\nDATA CDR TABLE    : ',op.d_cdr)

    # DATABASE CONNECTION
    vdb = pyodbc.connect(r'Driver={SQL Server};Server=%s;Database=%s;Trusted_Connection=yes;' %( op.srv, op.v_db ) )
    ddb = pyodbc.connect(r'Driver={SQL Server};Server=%s;Database=%s;Trusted_Connection=yes;' %( op.srv, op.d_db ) )
    vcur = vdb.cursor()
    dcur = ddb.cursor()

    print("DB CONNECTED, EXTRACTING CDRs...")
    cdr_voice = pd.read_sql("SELECT * FROM " + op.v_cdr, vdb)
    cdr_speech = pd.read_sql("SELECT * FROM " + op.s_cdr, vdb)
    cdr_data = pd.read_sql("SELECT * FROM " + op.d_cdr, ddb)
    vcur.close()
    dcur.close()
    vdb.close()
    ddb.close()
    print("Voice Records  :", len(cdr_voice))
    print("Speech Records :", len(cdr_speech))
    print("Data Records   :", len(cdr_data))

    print("DB CONNECTED, STARTING KPIS CALCULATION...")
    df = ck.buildKPIs(cdr_voice,cdr_speech,cdr_data,dc,dr,wc,wt)
    del cdr_voice
    del cdr_speech
    del cdr_data

    # WRITE TO EXCEL
    # writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')
    # df.to_excel(writer, sheet_name='KPI_Report')
    # writer.save()
    # SAVE TO DATABASE
    try:
        engine = sqlalchemy.create_engine(f'mssql+pyodbc://{op.srv}/{op.v_db}?driver=SQL+Server+Native+Client+11.0')
        df.to_sql(op.table_name, con = engine, if_exists='replace', sort=False)
        print("RESULTS WRITTEN TO DB: ", op.srv, ":", op.v_db, "TO TABLE: ", op.table_name)
        del engine
        del df
    except:
        print("writing to db failed")