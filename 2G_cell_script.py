# import Pandas
import pandas as pd

# import excel data and create a dataFrame
df = pd.read_excel('template_script.xlsx')

# create Dictionary  PARAMETER RELATED as key and the script as value
values = {
   'AMR':'SET GCELLCHMGAD:IDTYPE=BYNAME,CELLNAME="{}",AMRTCHHPRIORALLOW=ON,AMRTCHHPRIORLOAD={};{{{}}}',
'Idle SDCCH Threshold': 'SET GCELLCHMGBASIC:IDTYPE=BYNAME,CELLNAME="{}",IDLESDTHRES={},CELLMAXSD={};{{{}}}',
'CRO':'SET GCELLIDLEBASIC:IDTYPE=BYNAME,CELLNAME="{}",PI=YES,CRO={};{{{}}}',
'PT' : 'SET GCELLIDLEAD:IDTYPE=BYNAME,CELLNAME="{}",PT={};{{{}}}',
'Load Handover Support': 'SET GCELLHOBASIC:IDTYPE=BYNAME,CELLNAME="{}",LOADHOEN=YES;{{{}}}',
'Inter-layer HO Threshold' : 'SET GCELLHOBASIC:IDTYPE=BYNAME,CELLNAME="{}",HOTHRES={};{{{}}}',
'Min UL Level on Candidate Cell':'SET GCELLHOBASIC:IDTYPE=BYNAME,CELLNAME="{}",HOCDCMINDWPWR={},HOCDCMINUPPWR={};{{{}}}',
'T11':'SET GCELLTMR:IDTYPE=BYNAME,CELLNAME="{}",TIQUEUINGTIMER={};{{{}}}',
'FDD Qmin': 'SET GCELLCCUTRANSYS:IDTYPE=BYNAME,CELLNAME="{}",FDDQMIN={};{{{}}}',
'CRH' : 'SET GCELLIDLEBASIC:IDTYPE=BYNAME,CELLNAME="{}",CRH={};{{{}}}',
'MS MAX Retrans' : 'SET GCELLCCBASIC:IDTYPE=BYNAME,CELLNAME="{}",MSMAXRETRAN={};{{{}}}',
'TREESTABLISH' : 'SET GCELLTMR:IDTYPE=BYNAME,CELLNAME="{}",MSIPFAILINDDELAY={};{{{}}}',
'Minimum Access RXLEV' : 'SET GCELLBASICPARA:IDTYPE=BYNAME,CELLNAME="{}",RXMIN={};{{{}}}',
'Assignment Cell Load Judge Enable': 'SET GCELLCCBASIC:IDTYPE=BYNAME,CELLNAME="{}",ASSLOADJUDGEEN={};{{{}}}',
'CSRACH' : 'SET GCELLCCACCESS:IDTYPE=BYNAME,CELLNAME="{}",RACHACCLEV={};{{{}}}',
'T200 FACCH/F' : 'SET GCELLCCTMR:IDTYPE=BYNAME,CELLNAME="{}",T200FACCHF={},T200FACCHH={};{{{}}}',
'AFR SACCH Multi-Frames': 'SET GCELLCCBASIC:IDTYPE=BYNAME,CELLNAME="{}",AFRSAMULFRM={},AHRSAMULFRM={};{{{}}}',
'Signal Process After L2 Reestablished': 'SET GCELLSOFT:IDTYPE=BYNAME,CELLNAME="{}",L2REBSUCSIGPROCSW={};{{{}}}',
'Load HO Threshold':'SET GCELLHOAD:IDTYPE=BYNAME,CELLNAME="{}",TRIGTHRES={},LOADACCTHRES={};{{{}}}',
'Load HO Threshold_1':'SET GCELLHOAD:IDTYPE=BYNAME,CELLNAME="{}",TRIGTHRES={};{{{}}}',
'TCH Traffic Busy Threshold':'SET GCELLCHMGAD:IDTYPE=BYNAME,CELLNAME="{}",TCHBUSYTHRES={}; {{{}}}',
'Directed Retry': 'SET GCELLBASICPARA:IDTYPE=BYNAME,CELLNAME="{}",DIRECTRYEN={}; {{{}}}',
'Intracell F-H HO Allowed': 'SET GCELLHOBASIC:IDTYPE=BYNAME,CELLNAME="{}",HOCTRLSWITCH=HOALGORITHM1,INTRACELLFHHOEN={};',
'BQMARGIN' : 'MOD G2GNCELL:IDTYPE=BYNAME,SRC2GNCELLNAME="{}",NBR2GNCELLNAME="{}",NCELLTYPE=HANDOVERNCELL,SRCHOCTRLSWITCH=HOALGORITHM1,BQMARGIN={}; {{{}}}',
'MINOFFSET' : 'MOD G2GNCELL:IDTYPE=BYNAME,SRC2GNCELLNAME="{}",NBR2GNCELLNAME="{}",NCELLTYPE=HANDOVERNCELL,SRCHOCTRLSWITCH=HOALGORITHM1,MINOFFSET={}; {{{}}}',
'PBGTMARGIN' : 'MOD G2GNCELL:IDTYPE=BYNAME,SRC2GNCELLNAME="{}",NBR2GNCELLNAME="{}",NCELLTYPE=HANDOVERNCELL,SRCHOCTRLSWITCH=HOALGORITHM1,PBGTMARGIN={}; {{{}}}',
'Inter-cell HO Hysteresis' :'MOD G2GNCELL:IDTYPE=BYNAME,SRC2GNCELLNAME="{}",NBR2GNCELLNAME="{}",NCELLTYPE=HANDOVERNCELL,SRCHOCTRLSWITCH=HOALGORITHM1,INTERCELLHYST={}; {{{}}}',
'RMV G2GNCELL': 'RMV G2GNCELL:IDTYPE=BYNAME,SRC2GNCELLNAME="{}",NBR2GNCELLNAME="{}"; {{{}}}',
'LoadHOEn': 'SET GCELLHOBASIC:IDTYPE=BYNAME,CELLNAME="{}",LOADHOEN={};{{{}}}',
'Maximum Rate Threshold of PDCHs in a Cell' : 'SET GCELLPSCHM:IDTYPE=BYNAME,CELLNAME="{}",MAXPDCHRATE={}; {{{}}}',
'T3107': 'SET GCELLTMR:IDTYPE=BYNAME,CELLNAME="{}",ASSTIMER={}; {{{}}}',
'Activate L2 Re-establishment': 'SET GCELLSOFT: IDTYPE=BYNAME, CELLNAME="{}", ACTL2REEST={}; {{{}}}',
'Cell SDCCH Channel Maximum':'SET GCELLCHMGBASIC: IDTYPE=BYNAME, CELLNAME="{}", CELLMAXSD={}; {{{}}}',
'ADD G2GNCELL' : 'ADD G2GNCELL:IDTYPE=BYNAME,SRC2GNCELLNAME="{}",NBR2GNCELLNAME="{}",NCELLTYPE=HANDOVERNCELL,SRCHOCTRLSWITCH=HOALGORITHM1; {{{}}}',
'MAXTA' : 'SET GCELLBASICPARA: IDTYPE=BYNAME, CELLNAME="{}", MAXTA={}; {{{}}}',
'Deactivate TRX':'DEA GTRX:IDTYPE=BYID,TRXID={}; {{{}}}'


}

# iterate through the dataFrame

scripts_arr = [] * len(df.index)

for index, row in df.iterrows():
    rowIndex = df.index[index]
    cell_id = row['Cell ID']
    bsc_field = row['BSC/RNC']
    parameter_related = row['PARAMETER RELATED']
    proposed_value = row['Proposed Value']
    if 'AMR' in parameter_related:
        text = values['AMR']
        script = text.format(cell_id, proposed_value, bsc_field )
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'AMR TCH/H Prior Cell Load Threshold' in parameter_related:
        text = values['AMR']
        script = text.format(cell_id, proposed_value, bsc_field )
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Idle SDCCH Threshold' in parameter_related:

        if pd.isnull(row['Proposed Value.1']):
            text='SET GCELLCHMGBASIC:IDTYPE=BYNAME,CELLNAME="{}",IDLESDTHRES={};{{{}}}'
            script = text.format(cell_id, proposed_value,bsc_field)
        else:
            text = values['Idle SDCCH Threshold']
            script = text.format(cell_id, proposed_value, int(row['Proposed Value.1']), bsc_field)
        scripts_arr.append(script)
        df.loc[rowIndex, 'Scripts' ] = script
    elif 'CRO' in parameter_related:
        text = values['CRO']
        script= text.format(cell_id, proposed_value, bsc_field )
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Cell Reselect Offset' in parameter_related:
        text = values['CRO']
        script= text.format(cell_id, proposed_value, bsc_field )
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'PT' in parameter_related:
        text = values['PT']
        script = text.format(cell_id, proposed_value, bsc_field )
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Cell Reselect Penalty Time' in parameter_related:
        text = values['PT']
        script = text.format(cell_id, proposed_value, bsc_field )
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Load Handover Support'in parameter_related:
        text = values['Load Handover Support']
        script = text.format(cell_id, bsc_field )
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Inter-layer HO Threshold' in parameter_related:
        text = values['Inter-layer HO Threshold']
        script = text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Min UL Level on Candidate Cell' in parameter_related:
        text = values['Min UL Level on Candidate Cell']
        script = text.format(cell_id, int(row['Proposed Value.1']), proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'T11' in parameter_related:
        text = values['T11']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'FDD Qmin' in parameter_related:
        text = values['FDD Qmin']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'CRH' in parameter_related:
        text = values['CRH']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'MS MAX Retrans' in parameter_related:
        text = values['MS MAX Retrans']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'MS Max Retrans' in parameter_related:
        text = values['MS MAX Retrans']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)

    elif 'TREESTABLISH' in parameter_related:
        text = values['TREESTABLISH']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Minimum Access RXLEV' in parameter_related:
        text = values['Minimum Access RXLEV']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Assignment Cell Load Judge Enable' in parameter_related:
        text = values['Assignment Cell Load Judge Enable']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Assignment Cell Load Judge Enable' in parameter_related:
        text = values['Assignment Cell Load Judge Enable']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'CSRACH' in parameter_related:
        text = values['CSRACH']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'CS RACH Min. Access Level' in parameter_related:
        text = values['CSRACH']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'T200 FACCH/F' in parameter_related:
        text = values['T200 FACCH/F']
        script= text.format(cell_id,  proposed_value, int(row['Proposed Value.1']), bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'AFR SACCH Multi-Frames' in parameter_related:
        text = values['AFR SACCH Multi-Frames']
        script= text.format(cell_id,  proposed_value, int(row['Proposed Value.1']), bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Signal Process After L2 Reestablished' in parameter_related:
        text = values['Signal Process After L2 Reestablished']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'Load HO Threshold' in parameter_related:
        if pd.isnull(row['Proposed Value.1']):
            text = values['Load HO Threshold_1']
            script= text.format(cell_id, proposed_value,  bsc_field)
            df.loc[rowIndex, 'Scripts' ] = script
            scripts_arr.append(script)
        else:
            text = values['Load HO Threshold']
            script= text.format(cell_id, proposed_value, int(row['Proposed Value.1']),   bsc_field)
            df.loc[rowIndex, 'Scripts' ] = script
            scripts_arr.append(script)
    elif 'TCH Traffic Busy Threshold' in parameter_related:
        text = values['TCH Traffic Busy Threshold']
        script= text.format(cell_id,  proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Directed Retry' in parameter_related:
        text = values['Directed Retry']
        if  str(proposed_value).upper() == "YES":
            script= text.format(cell_id, "YES" , bsc_field)
        else:
            script= text.format(cell_id,  "NO" , bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'DR' in parameter_related:
        text = values['Directed Retry']
        if  str(proposed_value).upper() == "YES":
            script= text.format(cell_id, "YES" , bsc_field)
        else:
            script= text.format(cell_id, "NO" , bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Intracell F-H HO Allowed' in parameter_related:
        text = values['Intracell F-H HO Allowed']
        script= text.format(cell_id, proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'BQMARGIN' in parameter_related:
        text = values['BQMARGIN']
        script= text.format(cell_id,row['Site ID'], proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'BQ HO Margin' in parameter_related:
        text = values['BQMARGIN']
        script= text.format(cell_id,row['Site ID'], proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'MINOFFSET' in parameter_related:
        text = values['MINOFFSET']
        script= text.format(cell_id,row['Site ID'], proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'Min Access Level Offset' in parameter_related:
        text = values['MINOFFSET']
        script= text.format(cell_id,row['Site ID'], proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'PBGTMARGIN' in parameter_related:
        text = values['PBGTMARGIN']
        script= text.format(cell_id,row['Site ID'], proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'PBGT HO Threshold' in parameter_related:
        text = values['PBGTMARGIN']
        script= text.format(cell_id,row['Site ID'], proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'Inter-cell HO Hysteresis' in parameter_related:
        text = values['Inter-cell HO Hysteresis']
        script= text.format(cell_id,row['Site ID'], proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'INTERCELLHYST' in parameter_related:
        text = values['Inter-cell HO Hysteresis']
        script= text.format(cell_id,row['Site ID'], proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'RMV G2GNCELL' in parameter_related:
        text = values['RMV G2GNCELL']
        script= text.format(row['Current Value'],proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'LoadHOEn' in parameter_related:
        text = values['LoadHOEn']
        if  str(proposed_value).upper() == "YES":
            script= text.format(cell_id, "YES" , bsc_field)
        else:
            script= text.format(cell_id, "NO" , bsc_field)
        df.loc[rowIndex, 'Scripts' ] = script
        scripts_arr.append(script)
    elif 'Maximum Rate Threshold of PDCHs in a Cell' in parameter_related:
        text = values['Maximum Rate Threshold of PDCHs in a Cell']
        script= text.format(cell_id,proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif  'T3107' in parameter_related:
        text = values['T3107']
        script= text.format(cell_id,proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif  'Activate L2 Re-establishment' in parameter_related:
        text = values['Activate L2 Re-establishment']
        script= text.format(cell_id,proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif  'Cell SDCCH Channel Maximum' in parameter_related:
        text = values['Cell SDCCH Channel Maximum']
        script= text.format(cell_id,proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif 'ADD G2GNCELL' in parameter_related:
        text = values['ADD G2GNCELL']
        script = text.format(row['Current Value'],row['Proposed Value'], bsc_field )
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif  'MAXTA' in parameter_related:
        text = values['MAXTA']
        script= text.format(cell_id,proposed_value, bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)
    elif  'Deactivate TRX' in parameter_related:
        text = values['Deactivate TRX']
        script= text.format(int(row['Site ID']), bsc_field)
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)


    else:
        text= '__Parameter not available kindly contact Admin'
        script=text
        df.loc[rowIndex, 'Scripts'] = script
        scripts_arr.append(script)

df.to_excel('new_scripts_2.xlsx')
