import time
import numpy as np
import pandas as pd
import polars as pl
from copy import deepcopy
from pprint import pprint


def step1():
    """
    # In WIP_ARRANGEMENT.xlsx, Red col can direct copy paste to template_1,
    yellow col need to insert the formula in code like what we did in excel
    """
    df_wip_arrangement = pl.from_pandas(
        pd.read_excel(r"WIP_ARRANGEMENT.xlsx", sheet_name="DEV_DB")
    )
    print(df_wip_arrangement)

    new_dict = {
        "DEVICE": [],
        "BPC_WIP": [],
        "NPPH": [],
        "24HR": [],
        "Machine_Count": [],
    }
    for i in range(len(df_wip_arrangement)):
        new_dict["DEVICE"].append(df_wip_arrangement["DEVICE"][i])
        new_dict["BPC_WIP"].append(df_wip_arrangement["BPC_WIP"][i])
        new_dict["NPPH"].append(df_wip_arrangement["NPPH"][i])
        new_dict["24HR"].append(df_wip_arrangement["NPPH"][i] * 24)
        try:
            new_dict["Machine_Count"].append(
                df_wip_arrangement["WIP_5500"][i] / (df_wip_arrangement["NPPH"][i] * 24)
            )
        except ZeroDivisionError:
            new_dict["Machine_Count"].append(0)

    # pprint(new_dict)
    df_new = pl.from_dict(new_dict)
    df_new.write_excel("template_1.xlsx", table_name="DEV_DB")


def step2():
    """
    In BPCWIP excel, need to insert the new col called "PP_name", logic as follwoing
    (Firstexcel[WIRE] == BPCWIP[WIREPN],
     Firstexcel[CAPILLARY] == BPCWIP[CAPILLARY],
     Firstexcel[Device] == BPCWIP[Device],
     If all same, bring the Firstexcel[PP_NAME] in this col)
    """

    df_first_excel = pd.read_excel("First.xlsx")
    df_bpcwip = pd.read_excel("BPCWIP.xlsx")

    for i in range(len(df_bpcwip)):
        wirepn = df_bpcwip.loc[i, "WIREPN"]
        capillary = df_bpcwip.loc[i, "CAPILLARY"]
        device = df_bpcwip.loc[i, "DEVICE"]

        pp_name = df_first_excel[
            (df_first_excel["WIRE"] == wirepn)
            & (df_first_excel["CAPILLARY"] == capillary)
            & (df_first_excel["DEVICE"] == device)
        ]["PP_NAME"].values
        if len(pp_name) == 0:
            pp_name = ""
        else:
            pp_name = pp_name[0]
        df_bpcwip.loc[i, "PP_NAME"] = pp_name
    df_bpcwip = df_bpcwip.drop(columns=["PP_Name需要來自First excel"])
    df_bpcwip.to_excel("BPCWIP_2.xlsx", index=False)


def step3():
    output_data = {
        "LOC": [],
        "LOT": [],
        "EQ_NAME": [],
        "PP_NAME": [],
        "DEVICE": [],
        "NPPH": [],
        "Ranking:": [],
    }

    df_template_1 = pd.read_excel("template_1.xlsx")
    df_machine_setup = pd.read_excel("machine_setup.xlsx")
    df_bpcwip = pd.read_excel("BPCWIP_2.xlsx")
    df_bpcwip = deepcopy(df_bpcwip)[df_bpcwip["PP_NAME"].notnull()]

    print("-" * 25 + " template_1.xlsx column name " + "-" * 25)
    print(df_template_1.columns.tolist())
    print("-" * 25 + " machine_setup.xlsx column name " + "-" * 25)
    print(df_machine_setup.columns.tolist())
    print("-" * 25 + " BPCWIP_2.xlsx column name " + "-" * 25)
    print(df_bpcwip.columns.tolist())

    filter_condition = {"STATUS": "Non-RTD StandBy"}

    for i in range(len(df_machine_setup)):
        machine_setup_pp_name = df_machine_setup.loc[i, "PP"]
        eq_name = df_machine_setup.loc[i, "EQ_NAME"]
        device = df_machine_setup.loc[i, "DEVICE"]
        temp_material_data_df = df_bpcwip[
            (df_bpcwip["PP_NAME"] == machine_setup_pp_name)
            & (df_bpcwip["STATUS"] == filter_condition["STATUS"])
        ]
        if len(temp_material_data_df) > 2:
            dropped_index_list = temp_material_data_df.index.tolist()[:3]

            # take first 3 rows
            temp_material_data_df = temp_material_data_df[:3].reset_index(drop=True)

            # add to output_data
            for j in range(len(temp_material_data_df)):
                print(temp_material_data_df.loc[j, "LOC"])
                output_data["LOC"].append(temp_material_data_df.loc[j, "LOC"])
                output_data["LOT"].append(temp_material_data_df.loc[j, "LOT"])
                output_data["EQ_NAME"].append(eq_name)
                output_data["PP_NAME"].append(temp_material_data_df.loc[j, "PP_NAME"])
                output_data["DEVICE"].append(device)
                output_data["NPPH"].append(
                    df_template_1[df_template_1["DEVICE"] == device]["NPPH"].values[0]
                    if df_template_1[df_template_1["DEVICE"] == device]["NPPH"].values
                    else np.nan
                )
                output_data["Ranking:"].append(j + 1)

            # drop row by index list
            df_bpcwip = df_bpcwip.drop(dropped_index_list)
        elif len(temp_material_data_df) <= 2:
            dropped_index_list = temp_material_data_df.index.tolist()
            n = 3 - len(temp_material_data_df)

            # take first 3 rows
            temp_material_data_df = temp_material_data_df[:3].reset_index(drop=True)


    # df_machine_setup.to_excel("machine_setup_2.xlsx", index=False)
    df = pd.DataFrame(output_data)
    df.to_excel("output.xlsx", index=False)


if __name__ == "__main__":
    start_time = time.time()
    step3()
    print("--- %s seconds ---" % (time.time() - start_time))
