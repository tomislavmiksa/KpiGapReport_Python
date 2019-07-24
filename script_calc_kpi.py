# USED FOR DB ACCESS
# PYTHON DATA ANALYSIS LIBRARY
import pandas as pd
# WRITING TO EXCEL FILE
# local scripts and libs
import define_classes as kc
import math

def column_order():
    correct_order = ["kpi_levels", "grp_index",
                     "G_LEVEL_1", "G_LEVEL_2", "G_LEVEL_3", "G_LEVEL_4", "G_LEVEL_5", "TRAIN_INFO", "VENDOR",
                     "START_TIME", "END_TIME",
                     "CLASSIC_ATTEMPTS", "CLASSIC_COMPLETED", "CLASSIC_FAILED", "CLASSIC_DROPPED", "CLASSIC_CSSR", "CLASSIC_DCR", "CLASSIC_CSR", "CLASSIC_MIN_CST", "CLASSIC_AVG_CST", "CLASSIC_MAX_CST", "CLASSIC_P10_CST", "CLASSIC_P50_CST", "CLASSIC_P90_CST", "CLASSIC_BAD_CST", "CLASSIC_OKK_CST", "CLASSIC_BAD_CST_RATIO",
                     "VOLTE_CM", "VOLTE_CM_RATIO", "VOLTE_CM_END_RATIO", "CSFB_CM",
                     "WHATSAPP_ATTEMPTS", "WHATSAPP_COMPLETED", "WHATSAPP_FAILED", "WHATSAPP_DROPPED", "WHATSAPP_CSSR", "WHATSAPP_DCR", "WHATSAPP_CSR", "WHATSAPP_MIN_CST", "WHATSAPP_AVG_CST", "WHATSAPP_MAX_CST", "WHATSAPP_P10_CST", "WHATSAPP_P50_CST", "WHATSAPP_P90_CST", "WHATSAPP_BAD_CST", "WHATSAPP_OKK_CST", "WHATSAPP_BAD_CST_RATIO",
                     "CLASSIC_POLQA_ATTEMPTS", "CLASSIC_POLQA_BAD", "CLASSIC_POLQA_BAD_RATIO", "CLASSIC_POLQA_MIN_MOS", "CLASSIC_POLQA_AVG_MOS", "CLASSIC_POLQA_MAX_MOS", "CLASSIC_POLQA_P10_MOS", "CLASSIC_POLQA_P50_MOS", "CLASSIC_POLQA_P90_MOS", "CLASSIC_POLQA_VOLTE_MIN_MOS", "CLASSIC_POLQA_VOLTE_AVG_MOS", "CLASSIC_POLQA_VOLTE_MAX_MOS", "CLASSIC_POLQA_VOLTE_P10_MOS", "CLASSIC_POLQA_VOLTE_P50_MOS", "CLASSIC_POLQA_VOLTE_P90_MOS",
                     "WHATSAPP_POLQA_ATTEMPTS", "WHATSAPP_POLQA_BAD", "WHATSAPP_POLQA_BAD_RATIO", "WHATSAPP_POLQA_MIN_MOS", "WHATSAPP_POLQA_AVG_MOS", "WHATSAPP_POLQA_MAX_MOS", "WHATSAPP_POLQA_P10_MOS", "WHATSAPP_POLQA_P50_MOS", "WHATSAPP_POLQA_P90_MOS",
                     "HTTP_TRANSFER_LTE_RATIO", "HTTP_TRANSFER_LTE_UL_64QAM_RATIO", "HTTP_TRANSFER_LTE_DL_2564QAM RATIO",
                     "HTTP_TRANSFER_FDFS_DL_ATTEMPTS", "HTTP_TRANSFER_FDFS_DL_FAILED", "HTTP_TRANSFER_FDFS_DL_CUTOFF", "HTTP_TRANSFER_FDFS_DL_SUCCESS", "HTTP_TRANSFER_FDFS_DL_FAILED_RATIO", "HTTP_TRANSFER_FDFS_DL_CUTOFF_RATIO", "HTTP_TRANSFER_FDFS_DL_SUCCESS_RATIO",
                     "HTTP_TRANSFER_FDFS_UL_ATTEMPTS", "HTTP_TRANSFER_FDFS_UL_FAILED", "HTTP_TRANSFER_FDFS_UL_CUTOFF", "HTTP_TRANSFER_FDFS_UL_SUCCESS", "HTTP_TRANSFER_FDFS_UL_FAILED_RATIO", "HTTP_TRANSFER_FDFS_UL_CUTOFF_RATIO", "HTTP_TRANSFER_FDFS_UL_SUCCESS_RATIO", "HTTP_TRANSFER_FDTT_DL_ATTEMPTS", "HTTP_TRANSFER_FDTT_DL_FAILED", "HTTP_TRANSFER_FDTT_DL_CUTOFF", "HTTP_TRANSFER_FDTT_DL_SUCCESS", "HTTP_TRANSFER_FDTT_DL_MDR_MIN", "HTTP_TRANSFER_FDTT_DL_MDR_P10", "HTTP_TRANSFER_FDTT_DL_MDR_AVG", "HTTP_TRANSFER_FDTT_DL_MDR_P50", "HTTP_TRANSFER_FDTT_DL_MDR_P90", "HTTP_TRANSFER_FDTT_DL_MDR_MAX",
                     "HTTP_TRANSFER_FDTT_UL_ATTEMPTS" , "HTTP_TRANSFER_FDTT_UL_FAILED", "HTTP_TRANSFER_FDTT_UL_CUTOFF", "HTTP_TRANSFER_FDTT_UL_SUCCESS", "HTTP_TRANSFER_FDTT_UL_MDR_MIN", "HTTP_TRANSFER_FDTT_UL_MDR_P10", "HTTP_TRANSFER_FDTT_UL_MDR_AVG", "HTTP_TRANSFER_FDTT_UL_MDR_P50", "HTTP_TRANSFER_FDTT_UL_MDR_P90", "HTTP_TRANSFER_FDTT_UL_MDR_MAX", "HTTP_BROWSING_ATTEMPTS", "HTTP_BROWSING_FAILED", "HTTP_BROWSING_CUTOFF", "HTTP_BROWSING_SUCCESS", "HTTP_BROWSING_FAILED_RATIO", "HTTP_BROWSING_CUTOFF_RATIO", "HTTP_BROWSING_SUCCESS_RATIO", "HTTP_BROWSING_ROUNDTRIP_TIME_MIN", "HTTP_BROWSING_ROUNDTRIP_TIME_P10", "HTTP_BROWSING_ROUNDTRIP_TIME_AVG", "HTTP_BROWSING_ROUNDTRIP_TIME_P50", "HTTP_BROWSING_ROUNDTRIP_TIME_P90", "HTTP_BROWSING_ROUNDTRIP_TIME_MAX", "HTTP_BROWSING_CONTENT_TRANSFER_TIME_MIN", "HTTP_BROWSING_CONTENT_TRANSFER_TIME_P10", "HTTP_BROWSING_CONTENT_TRANSFER_TIME_AVG", "HTTP_BROWSING_CONTENT_TRANSFER_TIME_P50", "HTTP_BROWSING_CONTENT_TRANSFER_TIME_P90", "HTTP_BROWSING_CONTENT_TRANSFER_TIME_MAX",
                     "VIDEO_STREAM_ATTEMPTS", "VIDEO_STREAM_FAILED", "VIDEO_STREAM_CUTOFF", "VIDEO_STREAM_SUCCESS", "VIDEO_STREAM_IRRITATING_PLAYOUT", "VIDEO_STREAM_FAILED_RATIO", "VIDEO_STREAM_CUTOFF_RATIO", "VIDEO_STREAM_SUCCESS_RATIO", "VIDEO_STREAM_IRRITATING_PLAYOUT_RATIO", "VIDEO_STREAM_VMOS_MIN", "VIDEO_STREAM_VMOS_P10", "VIDEO_STREAM_VMOS_AVG", "VIDEO_STREAM_VMOS_P50", "VIDEO_STREAM_VMOS_P90", "VIDEO_STREAM_VMOS_MAX", "VIDEO_STREAM_VMOS_BAD", "VIDEO_STREAM_VMOS_OKK", "VIDEO_STREAM_VMOS_BAD_RATIO", "VIDEO_STREAM_TTFP_MIN", "VIDEO_STREAM_TTFP_P10", "VIDEO_STREAM_TTFP_AVG", "VIDEO_STREAM_TTFP_P50", "VIDEO_STREAM_TTFP_P90", "VIDEO_STREAM_TTFP_MAX", "VIDEO_STREAM_TTFP_BAD", "VIDEO_STREAM_TTFP_OKK", "VIDEO_STREAM_TTFP_BAD_RATIO",
                     "FACEBOOK_FDFS_UL_ATTEMPTS", "FACEBOOK_FDFS_UL_FAILED", "FACEBOOK_FDFS_UL_CUTOFF", "FACEBOOK_FDFS_UL_SUCCESS", "FACEBOOK_FDFS_UL_FAILED_RATIO", "FACEBOOK_FDFS_UL_CUTOFF_RATIO", "FACEBOOK_FDFS_UL_SUCCESS_RATIO", "FACEBOOK_FDFS_UL_TRANSFER_TIME_MIN", "FACEBOOK_FDFS_UL_TRANSFER_TIME_P10", "FACEBOOK_FDFS_UL_TRANSFER_TIME_AVG", "FACEBOOK_FDFS_UL_TRANSFER_TIME_P50", "FACEBOOK_FDFS_UL_TRANSFER_TIME_P90", "FACEBOOK_FDFS_UL_TRANSFER_TIME_MAX" ]
    return correct_order

def kpi_table_init():
    columns = column_order()
    return pd.DataFrame(columns=columns)

def kpi_report_kpis():
    columns = [  "Operator_Order"
                ,"Operator"
                ,"index"
                ,"kpi_levels"
                ,"grp_index"
                ,"G_LEVEL_1"
                ,"G_LEVEL_2"
                ,"G_LEVEL_3"
                ,"G_LEVEL_4"
                ,"G_LEVEL_5"
                ,"TRAIN_INFO"
                ,"VENDOR"
                ,"TIMEFRAME"
                ,"START_TIME"
                ,"END_TIME"
                ,"CLASSIC CALLS"
                ,"ATTEMPTS"
                ,"COMPLETED COUNT"
                ,"FAILED COUNT"
                ,"DROPPED COUNT"
                ,"CALL SETUP SUCCESS RATIO"
                ,"DROPPED CALL RATIO"
                ,"CALL SUCCESS RATIO"
                ,"MIN CALL SETUP TIME"
                ,"AVG CALL SETUP TIME"
                ,"MAX CALL SETUP TIME"
                ,"P10 CALL SETUP TIME"
                ,"P50 CALL SETUP TIME"
                ,"P90 CALL SETUP TIME"
                ,"CALL SETUP TIME >  15 s COUNT"
                ,"CALL SETUP TIME <= 15 s COUNT"
                ,"CALL SETUP TIME >  15 s RATIO"
                ,"VOLTE CALL MODE"
                ,"VOLTE CALL MODE RATIO"
                ,"VOLTE CALL MODE END RATIO"
                ,"CSFB CALL MODE RATIO"
                ,"WHATSAPP CALLS"
                ,"ATTEMPTS"
                ,"COMPLETED COUNT"
                ,"FAILED COUNT"
                ,"DROPPED COUNT"
                ,"CALL SETUP SUCCESS RATIO"
                ,"DROPPED CALL RATIO"
                ,"CALL SUCCESS RATIO"
                ,"MIN CALL SETUP TIME"
                ,"AVG CALL SETUP TIME"
                ,"MAX CALL SETUP TIME"
                ,"P10 CALL SETUP TIME"
                ,"P50 CALL SETUP TIME"
                ,"P90 CALL SETUP TIME"
                ,"CALL SETUP TIME >  15 s COUNT"
                ,"CALL SETUP TIME <= 15 s COUNT"
                ,"CALL SETUP TIME >  15 s RATIO"
                ,"CLASSIC CALLS SPEECH"
                ,"POLQA ATTEMPTS"
                ,"POLQA < 1.6 COUNT"
                ,"POLQA < 1.6 RATIO"
                ,"POLQA MIN MOS"
                ,"POLQA AVG MOS"
                ,"POLQA MAX MOS"
                ,"POLQA P10 MOS"
                ,"POLQA P50 MOS"
                ,"POLQA P90 MOS"
                ,"POLQA VOLTE MIN MOS"
                ,"POLQA VOLTE AVG MOS"
                ,"POLQA VOLTE MAX MOS"
                ,"POLQA VOLTE P10 MOS"
                ,"POLQA VOLTE P50 MOS"
                ,"POLQA VOLTE P90 MOS"
                ,"WHATSAPP CALLS SPEECH"
                ,"POLQA ATTEMPTS"
                ,"POLQA < 1.6 COUNT"
                ,"POLQA < 1.6 RATIO"
                ,"POLQA MIN MOS"
                ,"POLQA AVG MOS"
                ,"POLQA MAX MOS"
                ,"POLQA P10 MOS"
                ,"POLQA P50 MOS"
                ,"POLQA P90 MOS"
                ,"HTTP TRANSFER - TECHNOLOGY"
                ,"LTE Share"
                ,"UL64QAM used more than 30% in test session"
                ,"DL256QAM used more than 10% in test session"
                ,"HTTP TRANSFER - FDFS DL"
                ,"HTTP TRANSFER FDFS DL ATTEMPTS"
                ,"HTTP TRANSFER FDFS DL FAILED COUNT"
                ,"HTTP TRANSFER FDFS DL CUTOFF COUNT"
                ,"HTTP TRANSFER FDFS DL SUCCESS COUNT"
                ,"HTTP TRANSFER FDFS DL FAILED RATIO"
                ,"HTTP TRANSFER FDFS DL CUTOFF RATIO"
                ,"HTTP TRANSFER FDFS DL SUCCESS RATIO"
                ,"HTTP TRANSFER - FDFS DL"
                ,"HTTP TRANSFER FDFS UL ATTEMPTS"
                ,"HTTP TRANSFER FDFS UL FAILED COUNT"
                ,"HTTP TRANSFER FDFS UL CUTOFF COUNT"
                ,"HTTP TRANSFER FDFS UL SUCCESS COUNT"
                ,"HTTP TRANSFER FDFS UL FAILED RATIO"
                ,"HTTP TRANSFER FDFS UL CUTOFF RATIO"
                ,"HTTP TRANSFER FDFS UL SUCCESS RATIO"
                ,"HTTP TRANSFER - FDTT DL"
                ,"HTTP TRANSFER FDTT DL ATTEMPTS"
                ,"HTTP TRANSFER FDTT DL FAILED COUNT"
                ,"HTTP TRANSFER FDTT DL CUTOFF COUNT"
                ,"HTTP TRANSFER FDTT DL SUCCESS COUNT"
                ,"HTTP TRANSFER FDTT DL MDR MIN"
                ,"HTTP TRANSFER FDTT DL MDR P10"
                ,"HTTP TRANSFER FDTT DL MDR AVG"
                ,"HTTP TRANSFER FDTT DL MDR P50"
                ,"HTTP TRANSFER FDTT DL MDR P90"
                ,"HTTP TRANSFER FDTT DL MDR MAX"
                ,"HTTP TRANSFER - FDTT UL"
                ,"HTTP TRANSFER FDTT UL ATTEMPTS"
                ,"HTTP TRANSFER FDTT UL FAILED COUNT"
                ,"HTTP TRANSFER FDTT UL CUTOFF COUNT"
                ,"HTTP TRANSFER FDTT UL SUCCESS COUNT"
                ,"HTTP TRANSFER FDTT UL MDR MIN"
                ,"HTTP TRANSFER FDTT UL MDR P10"
                ,"HTTP TRANSFER FDTT UL MDR AVG"
                ,"HTTP TRANSFER FDTT UL MDR P50"
                ,"HTTP TRANSFER FDTT UL MDR P90"
                ,"HTTP TRANSFER FDTT UL MDR MAX"
                ,"HTTP BROWSING"
                ,"HTTP BROWSING ATTEMPTS"
                ,"HTTP BROWSING FAILED COUNT"
                ,"HTTP BROWSING CUTOFF COUNT"
                ,"HTTP BROWSING SUCCESS COUNT"
                ,"HTTP BROWSING FAILED RATIO"
                ,"HTTP BROWSING CUTOFF RATIO"
                ,"HTTP BROWSING SUCCESS RATIO"
                ,"HTTP BROWSING ROUNDTRIP TIME MIN"
                ,"HTTP BROWSING ROUNDTRIP TIME P10"
                ,"HTTP BROWSING ROUNDTRIP TIME AVG"
                ,"HTTP BROWSING ROUNDTRIP TIME P50"
                ,"HTTP BROWSING ROUNDTRIP TIME P90"
                ,"HTTP BROWSING ROUNDTRIP TIME MAX"
                ,"HTTP BROWSING CONTENT TRANSFER TIME MIN"
                ,"HTTP BROWSING CONTENT TRANSFER TIME P10"
                ,"HTTP BROWSING CONTENT TRANSFER TIME AVG"
                ,"HTTP BROWSING CONTENT TRANSFER TIME P50"
                ,"HTTP BROWSING CONTENT TRANSFER TIME P90"
                ,"HTTP BROWSING CONTENT TRANSFER TIME MAX"
                ,"VIDEO STREAM"
                ,"VIDEO STREAM ATTEMPTS"
                ,"VIDEO STREAM FAILED COUNT"
                ,"VIDEO STREAM CUTOFF COUNT"
                ,"VIDEO STREAM SUCCESS COUNT"
                ,"VIDEO STREAM IRRITATING PLAYOUT COUNT"
                ,"VIDEO STREAM FAILED RATIO"
                ,"VIDEO STREAM CUTOFF RATIO"
                ,"VIDEO STREAM SUCCESS RATIO"
                ,"VIDEO STREAM IRRITATING PLAYOUT RATIO"
                ,"VIDEO STREAM VMOS MIN"
                ,"VIDEO STREAM VMOS P10"
                ,"VIDEO STREAM VMOS AVG"
                ,"VIDEO STREAM VMOS P50"
                ,"VIDEO STREAM VMOS P90"
                ,"VIDEO STREAM VMOS MAX"
                ,"VIDEO STREAM VMOS < 1.6 COUNT"
                ,"VIDEO STREAM VMOS >= 1.6 COUNT"
                ,"VIDEO STREAM VMOS < 1.6 RATIO"
                ,"VIDEO STREAM TTFP MIN"
                ,"VIDEO STREAM TTFP P10"
                ,"VIDEO STREAM TTFP AVG"
                ,"VIDEO STREAM TTFP P50"
                ,"VIDEO STREAM TTFP P90"
                ,"VIDEO STREAM TTFP MAX"
                ,"VIDEO STREAM TTFP >= 10 s COUNT"
                ,"VIDEO STREAM TTFP < 10s COUNT"
                ,"VIDEO STREAM TTFP >= 10 s RATIO"
                ,"FACEBOOK - FDFS UL"
                ,"FACEBOOK FDFS UL ATTEMPTS"
                ,"FACEBOOK FDFS UL FAILED COUNT"
                ,"FACEBOOK FDFS UL CUTOFF COUNT"
                ,"FACEBOOK FDFS UL SUCCESS COUNT"
                ,"FACEBOOK FDFS UL FAILED RATIO"
                ,"FACEBOOK FDFS UL CUTOFF RATIO"
                ,"FACEBOOK FDFS UL SUCCESS RATIO"
                ,"FACEBOOK FDFS UL TRANSFER TIME MIN"
                ,"FACEBOOK FDFS UL TRANSFER TIME P10"
                ,"FACEBOOK FDFS UL TRANSFER TIME AVG"
                ,"FACEBOOK FDFS UL TRANSFER TIME P50"
                ,"FACEBOOK FDFS UL TRANSFER TIME P90"
                ,"FACEBOOK FDFS UL TRANSFER TIME MAX"]
    return columns

def append_kpis(glev, ind, ac, bc, cc):
    if cc.start_time is None:
        s_time = ac.start_time
    elif ac.start_time < cc.start_time:
        s_time = ac.start_time
    else:
        s_time = cc.start_time

    if ac.end_time is None:
        e_time = cc.end_time
    elif ac.end_time < cc.end_time:
        e_time = cc.end_time
    else:
        e_time = ac.end_time
    if e_time is None:
        e_time = s_time
    if s_time is None:
        s_time = e_time

    data = pd.DataFrame({"kpi_levels": glev,
                  "grp_index": ind,
                  "G_LEVEL_1": ac.g1,
                  "G_LEVEL_2": ac.g2,
                  "G_LEVEL_3": ac.g3,
                  "G_LEVEL_4": ac.g4,
                  "G_LEVEL_5": ac.g5,
                  "TRAIN_INFO": ac.train + '--' + ac.wagon,
                  "VENDOR": ac.vendor,
                  "START_TIME": s_time,
                  "END_TIME": e_time,
                  "CLASSIC_ATTEMPTS": ac.classic_attempts,
                  "CLASSIC_COMPLETED": ac.classic_completed,
                  "CLASSIC_FAILED": ac.classic_failed,
                  "CLASSIC_DROPPED": ac.classic_dropped,
                  "CLASSIC_CSSR": ac.classic_cssr,
                  "CLASSIC_DCR": ac.classic_dcr,
                  "CLASSIC_CSR": ac.classic_ccr,
                  "CLASSIC_MIN_CST": ac.classic_min_cst,
                  "CLASSIC_AVG_CST": ac.classic_avg_cst,
                  "CLASSIC_MAX_CST": ac.classic_max_cst,
                  "CLASSIC_P10_CST": ac.classic_p10_cst,
                  "CLASSIC_P50_CST": ac.classic_p50_cst,
                  "CLASSIC_P90_CST": ac.classic_p90_cst,
                  "CLASSIC_BAD_CST": ac.classic_poorCST,
                  "CLASSIC_OKK_CST": ac.classic_goodCST,
                  "CLASSIC_BAD_CST_RATIO": ac.classic_poorCSTratio,
                  "VOLTE_CM": ac.volte_start,
                  "VOLTE_CM_RATIO": ac.volte_ratio,
                  "VOLTE_CM_END_RATIO": ac.volte_end_ratio,
                  "CSFB_CM": ac.csfb_ratio,
                  "WHATSAPP_ATTEMPTS": ac.whatsapp_attempts,
                  "WHATSAPP_COMPLETED": ac.whatsapp_completed,
                  "WHATSAPP_FAILED": ac.whatsapp_failed,
                  "WHATSAPP_DROPPED": ac.whatsapp_dropped,
                  "WHATSAPP_CSSR": ac.whatsapp_cssr,
                  "WHATSAPP_DCR": ac.whatsapp_dcr,
                  "WHATSAPP_CSR": ac.whatsapp_ccr,
                  "WHATSAPP_MIN_CST": ac.whatsapp_min_cst,
                  "WHATSAPP_AVG_CST": ac.whatsapp_avg_cst,
                  "WHATSAPP_MAX_CST": ac.whatsapp_max_cst,
                  "WHATSAPP_P10_CST": ac.whatsapp_p10_cst,
                  "WHATSAPP_P50_CST": ac.whatsapp_p50_cst,
                  "WHATSAPP_P90_CST": ac.whatsapp_p90_cst,
                  "WHATSAPP_BAD_CST": ac.whatsapp_poorCST,
                  "WHATSAPP_OKK_CST": ac.whatsapp_goodCST,
                  "WHATSAPP_BAD_CST_RATIO": ac.whatsapp_poorCSTratio,
                  "CLASSIC_POLQA_ATTEMPTS": bc.classic_attempts,
                  "CLASSIC_POLQA_BAD": bc.classic_poorPMOS,
                  "CLASSIC_POLQA_BAD_RATIO": bc.classic_poorRMOS,
                  "CLASSIC_POLQA_MIN_MOS": bc.classic_min_mos,
                  "CLASSIC_POLQA_AVG_MOS": bc.classic_avg_mos,
                  "CLASSIC_POLQA_MAX_MOS": bc.classic_max_mos,
                  "CLASSIC_POLQA_P10_MOS": bc.classic_p10_mos,
                  "CLASSIC_POLQA_P50_MOS": bc.classic_p50_mos,
                  "CLASSIC_POLQA_P90_MOS": bc.classic_p90_mos,
                  "CLASSIC_POLQA_VOLTE_MIN_MOS": bc.classic_min_vmos,
                  "CLASSIC_POLQA_VOLTE_AVG_MOS": bc.classic_avg_vmos,
                  "CLASSIC_POLQA_VOLTE_MAX_MOS": bc.classic_max_vmos,
                  "CLASSIC_POLQA_VOLTE_P10_MOS": bc.classic_p10_vmos,
                  "CLASSIC_POLQA_VOLTE_P50_MOS": bc.classic_p50_vmos,
                  "CLASSIC_POLQA_VOLTE_P90_MOS": bc.classic_p90_vmos,
                  "WHATSAPP_POLQA_ATTEMPTS": bc.whatsapp_attempts,
                  "WHATSAPP_POLQA_BAD": bc.whatsapp_poorPMOS,
                  "WHATSAPP_POLQA_BAD_RATIO": bc.whatsapp_poorRMOS,
                  "WHATSAPP_POLQA_MIN_MOS": bc.whatsapp_min_mos,
                  "WHATSAPP_POLQA_AVG_MOS": bc.whatsapp_avg_mos,
                  "WHATSAPP_POLQA_MAX_MOS": bc.whatsapp_max_mos,
                  "WHATSAPP_POLQA_P10_MOS": bc.whatsapp_p10_mos,
                  "WHATSAPP_POLQA_P50_MOS": bc.whatsapp_p50_mos,
                  "WHATSAPP_POLQA_P90_MOS": bc.whatsapp_p90_mos,
                  "HTTP_TRANSFER_LTE_RATIO": cc.lte_share,
                  "HTTP_TRANSFER_LTE_UL_64QAM_RATIO": cc.ul_qam,
                  "HTTP_TRANSFER_LTE_DL_2564QAMRATIO": cc.dl_qam,
                  "HTTP_TRANSFER_FDFS_DL_ATTEMPTS": cc.fdfs_dl_attempts,
                  "HTTP_TRANSFER_FDFS_DL_FAILED": cc.fdfs_dl_failed,
                  "HTTP_TRANSFER_FDFS_DL_CUTOFF": cc.fdfs_dl_cutoff,
                  "HTTP_TRANSFER_FDFS_DL_SUCCESS": cc.fdfs_dl_succes,
                  "HTTP_TRANSFER_FDFS_DL_FAILED_RATIO": cc.fdfs_dl_ratio_failed,
                  "HTTP_TRANSFER_FDFS_DL_CUTOFF_RATIO": cc.fdfs_dl_ratio_cutoff,
                  "HTTP_TRANSFER_FDFS_DL_SUCCESS_RATIO": cc.fdfs_dl_ratio_succes,
                  "HTTP_TRANSFER_FDFS_UL_ATTEMPTS": cc.fdfs_ul_attempts,
                  "HTTP_TRANSFER_FDFS_UL_FAILED": cc.fdfs_ul_failed,
                  "HTTP_TRANSFER_FDFS_UL_CUTOFF": cc.fdfs_ul_cutoff,
                  "HTTP_TRANSFER_FDFS_UL_SUCCESS": cc.fdfs_ul_succes,
                  "HTTP_TRANSFER_FDFS_UL_FAILED_RATIO": cc.fdfs_ul_ratio_failed,
                  "HTTP_TRANSFER_FDFS_UL_CUTOFF_RATIO": cc.fdfs_ul_ratio_cutoff,
                  "HTTP_TRANSFER_FDFS_UL_SUCCESS_RATIO": cc.fdfs_ul_ratio_succes,
                  "HTTP_TRANSFER_FDTT_DL_ATTEMPTS": cc.fdtt_dl_attempts,
                  "HTTP_TRANSFER_FDTT_DL_FAILED": cc.fdtt_dl_failed,
                  "HTTP_TRANSFER_FDTT_DL_CUTOFF": cc.fdtt_dl_cutoff,
                  "HTTP_TRANSFER_FDTT_DL_SUCCESS": cc.fdtt_dl_succes,
                  "HTTP_TRANSFER_FDTT_DL_MDR_MIN": cc.fdtt_dl_min_mdr,
                  "HTTP_TRANSFER_FDTT_DL_MDR_P10": cc.fdtt_dl_p10_mdr,
                  "HTTP_TRANSFER_FDTT_DL_MDR_AVG": cc.fdtt_dl_avg_mdr,
                  "HTTP_TRANSFER_FDTT_DL_MDR_P50": cc.fdtt_dl_p50_mdr,
                  "HTTP_TRANSFER_FDTT_DL_MDR_P90": cc.fdtt_dl_p90_mdr,
                  "HTTP_TRANSFER_FDTT_DL_MDR_MAX": cc.fdtt_dl_max_mdr,
                  "HTTP_TRANSFER_FDTT_UL_ATTEMPTS": cc.fdtt_ul_attempts,
                  "HTTP_TRANSFER_FDTT_UL_FAILED": cc.fdtt_ul_failed,
                  "HTTP_TRANSFER_FDTT_UL_CUTOFF": cc.fdtt_ul_cutoff,
                  "HTTP_TRANSFER_FDTT_UL_SUCCESS": cc.fdtt_ul_succes,
                  "HTTP_TRANSFER_FDTT_UL_MDR_MIN": cc.fdtt_ul_min_mdr,
                  "HTTP_TRANSFER_FDTT_UL_MDR_P10": cc.fdtt_ul_p10_mdr,
                  "HTTP_TRANSFER_FDTT_UL_MDR_AVG": cc.fdtt_ul_avg_mdr,
                  "HTTP_TRANSFER_FDTT_UL_MDR_P50": cc.fdtt_ul_p50_mdr,
                  "HTTP_TRANSFER_FDTT_UL_MDR_P90": cc.fdtt_ul_p90_mdr,
                  "HTTP_TRANSFER_FDTT_UL_MDR_MAX": cc.fdtt_ul_max_mdr,
                  "HTTP_BROWSING_ATTEMPTS": cc.brws_dl_attempts,
                  "HTTP_BROWSING_FAILED": cc.brws_dl_failed,
                  "HTTP_BROWSING_CUTOFF": cc.brws_dl_cutoff,
                  "HTTP_BROWSING_SUCCESS": cc.brws_dl_succes,
                  "HTTP_BROWSING_FAILED_RATIO": cc.brws_dl_ratio_failed,
                  "HTTP_BROWSING_CUTOFF_RATIO": cc.brws_dl_ratio_cutoff,
                  "HTTP_BROWSING_SUCCESS_RATIO": cc.brws_dl_ratio_succes,
                  "HTTP_BROWSING_ROUNDTRIP_TIME_MIN": cc.brws_dl_min_rtt,
                  "HTTP_BROWSING_ROUNDTRIP_TIME_P10": cc.brws_dl_p10_rtt,
                  "HTTP_BROWSING_ROUNDTRIP_TIME_AVG": cc.brws_dl_avg_rtt,
                  "HTTP_BROWSING_ROUNDTRIP_TIME_P50": cc.brws_dl_p50_rtt,
                  "HTTP_BROWSING_ROUNDTRIP_TIME_P90": cc.brws_dl_p90_rtt,
                  "HTTP_BROWSING_ROUNDTRIP_TIME_MAX": cc.brws_dl_max_rtt,
                  "HTTP_BROWSING_CONTENT_TRANSFER_TIME_MIN": cc.brws_dl_min_ctt,
                  "HTTP_BROWSING_CONTENT_TRANSFER_TIME_P10": cc.brws_dl_p10_ctt,
                  "HTTP_BROWSING_CONTENT_TRANSFER_TIME_AVG": cc.brws_dl_avg_ctt,
                  "HTTP_BROWSING_CONTENT_TRANSFER_TIME_P50": cc.brws_dl_p50_ctt,
                  "HTTP_BROWSING_CONTENT_TRANSFER_TIME_P90": cc.brws_dl_p90_ctt,
                  "HTTP_BROWSING_CONTENT_TRANSFER_TIME_MAX": cc.brws_dl_max_ctt,
                  "VIDEO_STREAM_ATTEMPTS": cc.vs_yt_dl_attempts,
                  "VIDEO_STREAM_FAILED": cc.vs_yt_dl_failed,
                  "VIDEO_STREAM_CUTOFF": cc.vs_yt_dl_cutoff,
                  "VIDEO_STREAM_SUCCESS": cc.vs_yt_dl_succes,
                  "VIDEO_STREAM_IRRITATING_PLAYOUT": cc.vs_yt_dl_irritating,
                  "VIDEO_STREAM_FAILED_RATIO": cc.vs_yt_dl_ratio_failed,
                  "VIDEO_STREAM_CUTOFF_RATIO": cc.vs_yt_dl_ratio_cutoff,
                  "VIDEO_STREAM_SUCCESS_RATIO": cc.vs_yt_dl_ratio_succes,
                  "VIDEO_STREAM_IRRITATING_PLAYOUT_RATIO" : cc.vs_yt_dl_ratio_irritating,
                  "VIDEO_STREAM_VMOS_MIN": cc.vs_yt_dl_min_vMOS,
                  "VIDEO_STREAM_VMOS_P10": cc.vs_yt_dl_p10_vMOS,
                  "VIDEO_STREAM_VMOS_AVG": cc.vs_yt_dl_avg_vMOS,
                  "VIDEO_STREAM_VMOS_P50": cc.vs_yt_dl_p50_vMOS,
                  "VIDEO_STREAM_VMOS_P90": cc.vs_yt_dl_p90_vMOS,
                  "VIDEO_STREAM_VMOS_MAX": cc.vs_yt_dl_max_vMOS,
                  "VIDEO_STREAM_VMOS_BAD": cc.vs_yt_dl_poor_vMOS,
                  "VIDEO_STREAM_VMOS_OKK": cc.vs_yt_dl_good_vMOS,
                  "VIDEO_STREAM_VMOS_BAD_RATIO": cc.vs_yt_dl_prat_vMOS,
                  "VIDEO_STREAM_TTFP_MIN": cc.vs_yt_dl_min_ttfp,
                  "VIDEO_STREAM_TTFP_P10": cc.vs_yt_dl_p10_ttfp,
                  "VIDEO_STREAM_TTFP_AVG": cc.vs_yt_dl_avg_ttfp,
                  "VIDEO_STREAM_TTFP_P50": cc.vs_yt_dl_p50_ttfp,
                  "VIDEO_STREAM_TTFP_P90": cc.vs_yt_dl_p90_ttfp,
                  "VIDEO_STREAM_TTFP_MAX": cc.vs_yt_dl_max_ttfp,
                  "VIDEO_STREAM_TTFP_BAD": cc.vs_yt_dl_poor_ttfp,
                  "VIDEO_STREAM_TTFP_OKK": cc.vs_yt_dl_good_ttfp,
                  "VIDEO_STREAM_TTFP_BAD_RATIO": cc.vs_yt_dl_prat_ttfp,
                  "FACEBOOK_FDFS_UL_ATTEMPTS": cc.fb_fdfs_ul_attempts,
                  "FACEBOOK_FDFS_UL_FAILED": cc.fb_fdfs_ul_failed,
                  "FACEBOOK_FDFS_UL_CUTOFF": cc.fb_fdfs_ul_cutoff,
                  "FACEBOOK_FDFS_UL_SUCCESS": cc.fb_fdfs_ul_succes,
                  "FACEBOOK_FDFS_UL_FAILED_RATIO": cc.fb_fdfs_ul_ratio_failed,
                  "FACEBOOK_FDFS_UL_CUTOFF_RATIO": cc.fb_fdfs_ul_ratio_cutoff,
                  "FACEBOOK_FDFS_UL_SUCCESS_RATIO": cc.fb_fdfs_ul_ratio_succes,
                  "FACEBOOK_FDFS_UL_TRANSFER_TIME_MIN": cc.fb_fdfs_ul_min_ctt,
                  "FACEBOOK_FDFS_UL_TRANSFER_TIME_P10": cc.fb_fdfs_ul_p10_ctt,
                  "FACEBOOK_FDFS_UL_TRANSFER_TIME_AVG": cc.fb_fdfs_ul_avg_ctt,
                  "FACEBOOK_FDFS_UL_TRANSFER_TIME_P50": cc.fb_fdfs_ul_p50_ctt,
                  "FACEBOOK_FDFS_UL_TRANSFER_TIME_P90": cc.fb_fdfs_ul_p90_ctt,
                  "FACEBOOK_FDFS_UL_TRANSFER_TIME_MAX": cc.fb_fdfs_ul_max_ctt}, index = [0])
    return data


def buildKPIs(cdr_v,cdr_s,cdr_d,m1,m2,m3,m4):
    # LEVEL G0
    df = kpi_table_init()
    a = kc.Voice(cdr_v,
                g1 = '',
                g2 = '',
                g3 = '',
                g4 = '',
                g5 = '',
                vend = '',
                reg = '',
                t_name = '',
                w_name = '')
    b = kc.Speech(cdr_s,
                g1 = '',
                g2 = '',
                g3 = '',
                g4 = '',
                g5 = '',
                vend = '',
                reg = '',
                t_name = '',
                w_name = '')
    c = kc.Data(cdr_d,
                g1 = '',
                g2 = '',
                g3 = '',
                g4 = '',
                g5 = '',
                vend = '',
                reg = '',
                t_name = '',
                w_name = '')
    d1 = append_kpis('g0', 0, a, b, c)
    df = df.append(d1)
    del d1
    # LEVEL G2
    i = 0
    for module in [m1,m2,m3,m4]:
        i = i + 1
        a = kc.Voice(cdr_v,
                    g1 = module.g1,
                    g2 = module.g2,
                    g3 = '',
                    g4 = '',
                    g5 = '',
                    vend = '',
                    reg = '',
                    t_name = '',
                    w_name = '')
        b = kc.Speech(cdr_s,
                    g1 = module.g1,
                    g2 = module.g2,
                    g3 = '',
                    g4 = '',
                    g5 = '',
                    vend = '',
                    reg = '',
                    t_name = '',
                    w_name = '')
        c = kc.Data(cdr_d,
                    g1 = module.g1,
                    g2 = module.g2,
                    g3 = '',
                    g4 = '',
                    g5 = '',
                    vend = '',
                    reg = '',
                    t_name = '',
                    w_name = '')
        d1 = append_kpis('g2', i, a, b, c)
        df = df.append(d1)
        del d1

        # G_Level_2 + Vendor
        j = 0
        for sub in module.g2v:
            tvend = sub.split(';')[1]
            if tvend != "":
                j = j+1
                a = kc.Voice(cdr_v,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = '',
                            g4 = '',
                            g5 = '',
                            vend = tvend,
                            reg = '',
                            t_name = '',
                            w_name = '')
                b = kc.Speech(cdr_s,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = '',
                            g4 = '',
                            g5 = '',
                            vend = tvend,
                            reg = '',
                            t_name = '',
                            w_name = '')
                c = kc.Data(cdr_d,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = '',
                            g4 = '',
                            g5 = '',
                            vend = tvend,
                            reg = '',
                            t_name = '',
                            w_name = '')
                d1 = append_kpis('g2v', i*10+j, a, b, c)
                df = df.append(d1)
                del d1

        # G_Level_3
        j = 0
        for sub in module.g3:
            if sub != "":
                j = j+1
                a = kc.Voice(cdr_v,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = sub,
                            g4 = '',
                            g5 = '',
                            vend = '',
                            reg = '',
                            t_name = '',
                            w_name = '')
                b = kc.Speech(cdr_s,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = sub,
                            g4 = '',
                            g5 = '',
                            vend = '',
                            reg = '',
                            t_name = '',
                            w_name = '')
                c = kc.Data(cdr_d,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = sub,
                            g4 = '',
                            g5 = '',
                            vend = '',
                            reg = '',
                            t_name = '',
                            w_name = '')
                d1 = append_kpis('g3', i*10+j, a, b, c)
                df = df.append(d1)
                del d1

        # G_Level_3 + Vendor
        j = 0
        for sub in module.g3v:
            if sub.split(';')[1] != "":
                j = j+1
                a = kc.Voice(cdr_v,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = sub.split(';')[0],
                            g4 = '',
                            g5 = '',
                            vend = sub.split(';')[1],
                            reg = '',
                            t_name = '',
                            w_name = '')
                b = kc.Speech(cdr_s,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = sub.split(';')[0],
                            g4 = '',
                            g5 = '',
                            vend = sub.split(';')[1],
                            reg = '',
                            t_name = '',
                            w_name = '')
                c = kc.Data(cdr_d,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = sub.split(';')[0],
                            g4 = '',
                            g5 = '',
                            vend = sub.split(';')[1],
                            reg = '',
                            t_name = '',
                            w_name = '')
                d1 = append_kpis('g3v', i*10+j, a, b, c)
                df = df.append(d1)
                del d1

        # G_Level_4
        j = 0
        if module.g2 != 'Train Route':
            for sub in module.g4:
                if sub.split(';')[1] != "":
                    j = j+1
                    a = kc.Voice(cdr_v,
                                g1 = module.g1,
                                g2 = module.g2,
                                g3 = sub.split(';')[0],
                                g4 =  sub.split(';')[1],
                                g5 = '',
                                vend = '',
                                reg = '',
                                t_name = '',
                                w_name = '')
                    b = kc.Speech(cdr_s,
                                g1 = module.g1,
                                g2 = module.g2,
                                g3 = sub.split(';')[0],
                                g4 =  sub.split(';')[1],
                                g5 = '',
                                vend = '',
                                reg = '',
                                t_name = '',
                                w_name = '')
                    c = kc.Data(cdr_d,
                                g1 = module.g1,
                                g2 = module.g2,
                                g3 = sub.split(';')[0],
                                g4 =  sub.split(';')[1],
                                g5 = '',
                                vend = '',
                                reg = '',
                                t_name = '',
                                w_name = '')
                    d1 = append_kpis('g4', i*100+j, a, b, c)
                    df = df.append(d1)
                    del d1
        else:
            for sub in module.tn:
                if sub.split(';')[1] != "":
                    j = j+1
                    a = kc.Voice(cdr_v,
                                g1 = module.g1,
                                g2 = module.g2,
                                g3 = '',
                                g4 =  sub.split(';')[0],
                                g5 = '',
                                vend = '',
                                reg = '',
                                t_name = sub.split(';')[1],
                                w_name = sub.split(';')[2])
                    b = kc.Speech(cdr_s,
                                g1 = module.g1,
                                g2 = module.g2,
                                g3 = '',
                                g4 =  sub.split(';')[0],
                                g5 = '',
                                vend = '',
                                reg = '',
                                t_name = sub.split(';')[1],
                                w_name = sub.split(';')[2])
                    c = kc.Data(cdr_d,
                                g1 = module.g1,
                                g2 = module.g2,
                                g3 = '',
                                g4 =  sub.split(';')[0],
                                g5 = '',
                                vend = '',
                                reg = '',
                                t_name = sub.split(';')[1],
                                w_name = sub.split(';')[2])
                    d1 = append_kpis('g4', i*100+j, a, b, c)
                    df = df.append(d1)
                    del d1
        # G_Level_5
        j = 0
        for sub in module.g5:
            if sub != "":
                j = j+1
                a = kc.Voice(cdr_v,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = '',
                            g4 = sub.split(';')[0],
                            g5 = sub.split(';')[1],
                            vend = '',
                            reg = '',
                            t_name = '',
                            w_name = '')
                b = kc.Speech(cdr_s,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = '',
                            g4 = sub.split(';')[0],
                            g5 = sub.split(';')[1],
                            vend = '',
                            reg = '',
                            t_name = '',
                            w_name = '')
                c = kc.Data(cdr_d,
                            g1 = module.g1,
                            g2 = module.g2,
                            g3 = '',
                            g4 = sub.split(';')[0],
                            g5 = sub.split(';')[1],
                            vend = '',
                            reg = '',
                            t_name = '',
                            w_name = '')
                d1 = append_kpis('g5', i*1000+j, a, b, c)
                df = df.append(d1)
                del d1
    return df

def write_excel(reports, group, writer, sheet, g_lev_1='', g_lev_2='',ind=[], col = []):
    worksheet = writer.add_worksheet(sheet)

    # KPI NAMES
    lheader_fmt = writer.add_format({'align': 'left',
                                     'font_size': 8})
    lheader_fmt.set_bold()
    lheader_fmt.set_font_color('white')
    lheader_fmt.set_bg_color('#003366')
    lheader_fmt.set_border()

    # BOLD KPI NAMES
    blheader_fmt = writer.add_format({'align': 'left',
                                      'font_size': 11,
                                      'bold': True})
    blheader_fmt.set_font_color('white')
    blheader_fmt.set_bg_color('#003366')
    blheader_fmt.set_border()

    # HEADER WITH TESTS NAMES
    mheader_fmt = writer.add_format({'align': 'left',
                                     'font_size': 8})
    mheader_fmt.set_bold()
    mheader_fmt.set_italic()
    mheader_fmt.set_font_color('#003366')
    mheader_fmt.set_font_size(11)

    # OTHERS
    normal_fmt  = writer.add_format({'align': 'center',
                                     'font_size': 8})
    normal_fmt.set_border()
    # OTHERS BOLD
    bnormal_fmt  = writer.add_format({'align': 'center',
                                     'font_size': 11,
                                     'bold': True})
    bnormal_fmt.set_border()
    # NORMAL PERCENTS
    percent_fmt = writer.add_format({'num_format': '0.00%',
                                     'align': 'center',
                                     'font_size': 8})
    percent_fmt.set_border()
    # BOLD PERCENTS
    bpercent_fmt = writer.add_format({'num_format': '0.00%',
                                     'align': 'center',
                                     'font_size': 11,
                                      'bold': True})
    bpercent_fmt.set_border()
    # NORMAL DATE
    date_fmt    = writer.add_format({'num_format': 'mm/dd/yyyy',
                                     'align': 'center',
                                     'font_size': 11})
    date_fmt.set_bold()
    date_fmt.set_italic()
    date_fmt.set_border()
    # NORMAL SECONDS
    seconds_fmt = writer.add_format({'num_format': '0.00" [s]"',
                                     'align': 'center',
                                     'font_size': 8})
    seconds_fmt.set_border()
    # BOLD SECONDS
    bseconds_fmt = writer.add_format({'num_format': '0.00" [s]"',
                                     'align': 'center',
                                     'font_size': 11,
                                      'bold': True})
    bseconds_fmt.set_border()
    # NORMAL MILISECONDS
    mseconds_fmt= writer.add_format({'num_format': '0" [ms]"',
                                     'align': 'center',
                                     'font_size': 8})
    mseconds_fmt.set_border()
    # BOLD MILISECONDS
    bmseconds_fmt= writer.add_format({'num_format': '0" [ms]"',
                                     'align': 'center',
                                     'font_size': 11,
                                     'bold': True})
    bmseconds_fmt.set_border()
    # NORMAL MOS
    mos_fmt     = writer.add_format({'num_format': '0.00" [MOS]"',
                                     'align': 'center',
                                     'font_size': 8})
    mos_fmt.set_border()
    # BOLD MOS
    bmos_fmt     = writer.add_format({'num_format': '0.00" [MOS]"',
                                     'align': 'center',
                                     'font_size': 11,
                                     'bold': True})
    bmos_fmt.set_border()
    # NORMAL Mbit/s
    mbit_fmt    = writer.add_format({'num_format': '0.00" [Mbit/s]"',
                                     'align': 'center',
                                     'font_size': 8})
    mbit_fmt.set_border()
    # BOLD Mbit/s
    bmbit_fmt    = writer.add_format({'num_format': '0.00" [Mbit/s]"',
                                     'align': 'center',
                                     'font_size': 11,
                                     'bold': True})
    bmbit_fmt.set_border()

    for view in reports:
        tmp = view[view.kpi_levels == group]
        if len(ind) > 0:
            tmp = tmp[tmp['Operator_Order'].isin(ind)]
        if g_lev_1 != '':
            tmp = tmp[tmp.G_LEVEL_1 == g_lev_1]
        if g_lev_2 != '':
            tmp = tmp[tmp.G_LEVEL_2 == g_lev_2]
        if 'output' in locals():
            output = output.append(tmp)
        else:
            output = tmp
    output = output.sort_values(by=['grp_index', 'Operator_Order'])
    i = 0
    for value in list(output):
        j = 1
        for v in list(output[value]):
            try:
                worksheet.write(i, j, v)
                if i in [13, 14]:
                    worksheet.set_row(i, i, date_fmt)
                elif i in [20,21,22,31,33,34,35,41,42,43,52,56,72,80,90,98,128,149,150,159,168,176]:
                    worksheet.set_row(i, i, bpercent_fmt)
                elif i in [81, 82, 88, 89, 96, 97, 126, 127, 147, 148, 168, 174, 175]:
                    worksheet.set_row(i, i, percent_fmt)
                elif i in [24,45,137,162,179]:
                    worksheet.set_row(i, i, bseconds_fmt)
                elif i in [23, 25, 26, 27, 28, 44, 46, 47, 48, 49, 135, 136,137, 138, 139, 140, 149, 160, 161, 163, 164, 177, 178, 180, 181, 182]:
                    worksheet.set_row(i, i, seconds_fmt)
                elif i in [58,64,74,153]:
                    worksheet.set_row(i, i, bmos_fmt)
                elif i in [57, 58, 59, 60, 61, 62, 63, 65, 66, 67, 68, 73, 75, 76, 77, 78, 149, 151, 152, 154, 155, 156]:
                    worksheet.set_row(i, i, mos_fmt)
                elif i in [131]:
                    worksheet.set_row(i, i, bmseconds_fmt)
                elif i in [129, 130, 132, 133, 134]:
                    worksheet.set_row(i, i, mseconds_fmt)
                elif i in [105, 106, 108, 116, 117, 119]:
                    worksheet.set_row(i, i, bmbit_fmt)
                elif i in [104, 107, 109, 115, 118, 120]:
                    worksheet.set_row(i, i, mbit_fmt)
                elif i in [5,6,7,8,9,10,11]:
                    worksheet.set_row(i, i, bnormal_fmt)
                else:
                    worksheet.set_row(i, i, normal_fmt)
            except:
                pass
            j += 1
        worksheet.write(i, 0, value, lheader_fmt)
        i += 1
        # SKIP NAMES OF THE TEST WHICH DO NOT EXISTS IN DATABASE
        if i in [12, 15, 36, 53, 69, 79, 83, 91, 99, 110, 121, 141, 169]:
            i += 1

    del output
    f = 0
    for v in kpi_report_kpis():
        if f in [12, 15, 36, 53, 69, 79, 83, 91, 99, 110, 121, 141, 169]:
            worksheet.write(f, 0, v, mheader_fmt)
        elif f in [20,21,22,31,33,34,35,41,42,43,52,56,72,80,90,98,128,149,150,159,168,176,24,45,137,162,179,58,64,74,153,131,105, 106, 108, 116, 117, 119,5,6,7,8,9,10,11]:
            worksheet.write(f, 0, v, blheader_fmt)
        else:
            worksheet.write(f, 0, v, lheader_fmt)
        f += 1

    # FREEZE 1st 15 ROWS
    worksheet.freeze_panes(15, 0)

    # worksheet.set_row(0, None, None, {'level': 1, 'collapsed': True})
    # worksheet.set_row(0, None, None, {'level': 3, 'collapsed': True})
    # worksheet.set_default_row(hide_unused_rows=True)

def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string