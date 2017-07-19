Automated FM Data Validation

To run these programs, you will need to install openpyxl to access data from excel sheets. Use this command:
sudo pip install openpyxl

To run programs from the command line, use the following form for all overlay processing (replace the <> with file names):

python <script_name> <master_data_sheet_file> <overlay_file>

For the special FM_Validation, type in:

python fm_validation_script.py FM_Data.xlsx FM_Overlay.xlsx


Note: The objectives below are for the FM_Data overlay, but the exact same methodology is employed for all other overlays

Objective 1: To compare the FISCAM controls from the master FM Document (located in "FM_Data.xlsx", column C)
to the actual data in the FM_Overlay prepared by Matrix (located in "FM_Overlay.xlsx", column C) and note any differenes (misssing/erroneous requirements).

I will take the following steps to check for missing requirements:

1) Load data sheet and overlay sheet
2) Clean input data and store all control requirements (column C in excel doc) from data sheet in reqs_dict (in populate_req_dict_function)
3) Clean input data and store all controls present in overlay data (column C in excel doc) in data_dict (in populate_data_dict function)
4) Cross-reference each entry in reqs_dict with data_dict to look for missing requirements and write them to output file (in compare_reqs function)
5) Cross-reference each entry in data_dict with reqs_ddict to look for erroneous additions and write them to output file (in compare_reqs function)

Objective 2: To compare the the Control Ranking from the master FM document(located in "FM_Data.xlsx", column C)
to the actual rankings in the FM_Overlay prepared by Matrix (located in "FM_Overlay.xlsx", column C) and note any differences.

I will take the following steps:

1) Load data and overlay sheet
2) Store labels from data sheet (located in "FM_Data.xlsx", column D) in rank_master_dict (in pop_master_rank_reqs function)
3) Store labels from FM_Overlay prepared by Matrix (located in "FM_Overlay.xlsx", column D) in rank_data_dict (in pop_data_rank function)
     3a) Every sub-point in the overlay should have the same label as the master says. For example, the label associated with
         'AC-2' in the data sheet ("FM_Data.xlsx, column D) is "Primary". As such, every instance of AC-2 in the FM_Overlay
         , and every sub-point (AC-2(2), AC-2(3), etc.) should also be associated with the label "Primary"
4) Cross-reference each entry in the FM_Overlay with the data sheet to ensure label consistency and write any inconsistencies
   to output file (in compare_ranks function)
