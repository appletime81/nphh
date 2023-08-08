import time
import numpy as np
import pandas as pd
import polars as pl
from copy import deepcopy
from pprint import pprint


def step1():
    """
    # In WIP_ARRANGEMENT.xlsx, Red col can direct copy and paste to template_1,
    yellow col need to insert the formula in code like what we did in Excel
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
        "Ranking": [],
    }
    filter_condition = {"STATUS": "Non-RTD StandBy"}

    # ---------------------- Second.xlsx ----------------------
    df_second = pd.read_excel("Second.xlsx")
    df_second = df_second.sort_values(by=["NPPH"], ascending=False).reset_index(
        drop=True
    )  # sorted by nphh

    # ---------------------- BPCWIP_2.xlsx ----------------------
    df_bpcwip = pd.read_excel("BPCWIP_2.xlsx")
    df_bpcwip = deepcopy(df_bpcwip)[df_bpcwip["PP_NAME"].notnull()]

    # ---------------------- machine_setup.xlsx ----------------------
    df_machine_setup = pd.read_excel("machine_setup.xlsx")

    for i in range(len(df_machine_setup)):
        machine_setup_pp_name = df_machine_setup.loc[i, "PP"]
        eq_name = df_machine_setup.loc[i, "EQ_NAME"]
        device = df_machine_setup.loc[i, "DEVICE"]
        temp_material_data_df = deepcopy(
            df_bpcwip[
                (df_bpcwip["PP_NAME"] == machine_setup_pp_name)
                & (df_bpcwip["STATUS"] == filter_condition["STATUS"])
            ]
        )
        if len(temp_material_data_df) > 2:
            dropped_index_list = temp_material_data_df.index.tolist()[:3]

            # take first 3 rows
            temp_material_data_df = temp_material_data_df[:3].reset_index(drop=True)

            # add to output_data
            for j in range(len(temp_material_data_df)):
                # print(temp_material_data_df.loc[j, "LOC"])
                output_data["LOC"].append(temp_material_data_df.loc[j, "LOC"])
                output_data["LOT"].append(temp_material_data_df.loc[j, "LOT"])
                output_data["EQ_NAME"].append(eq_name)
                output_data["PP_NAME"].append(temp_material_data_df.loc[j, "PP_NAME"])
                output_data["DEVICE"].append(device)
                output_data["NPPH"].append(
                    df_second[df_second["DEVICE"] == device]["NPPH"].values[0]
                    if len(df_second[df_second["DEVICE"] == device]["NPPH"].values)
                    else np.nan
                )
                output_data["Ranking"].append(j + 1)

            # drop row by index list
            df_bpcwip = df_bpcwip.drop(dropped_index_list)
        elif len(temp_material_data_df) <= 2:
            n = 3 - len(temp_material_data_df)
            count = 1
            # ------------------------ take first 3(<=3) rows ------------------------
            # get first len(temp_material_data_df) rows
            if len(temp_material_data_df) > 0:
                dropped_index_list = temp_material_data_df.index.tolist()
                temp_material_data_df = temp_material_data_df.reset_index(drop=True)
                for j in range(len(temp_material_data_df)):
                    # print(temp_material_data_df.loc[j, "LOC"])
                    output_data["LOC"].append(temp_material_data_df.loc[j, "LOC"])
                    output_data["LOT"].append(temp_material_data_df.loc[j, "LOT"])
                    output_data["EQ_NAME"].append(eq_name)
                    output_data["PP_NAME"].append(
                        temp_material_data_df.loc[j, "PP_NAME"]
                    )
                    output_data["DEVICE"].append(device)
                    output_data["NPPH"].append(
                        df_second[df_second["DEVICE"] == device][
                            "NPPH"
                        ].values.tolist()[0]
                        if df_second[df_second["DEVICE"] == device][
                            "NPPH"
                        ].values.tolist()
                        else np.nan
                    )
                    output_data["Ranking"].append(count)
                    count += 1
                df_bpcwip = df_bpcwip.drop(dropped_index_list)

            # ------------------------ get first three n from Second.xlsx ------------------------
            temp_second_df = deepcopy(df_second)
            temp_second_df = temp_second_df[
                (temp_second_df["MACHINE_NAME"] == eq_name)
                & (temp_second_df["DEVICE"] == device)
            ]
            temp_second_df_PP_NAME_list = temp_second_df["PP_NAME"].values.tolist()
            temp_material_data_df_for_second_df = deepcopy(
                df_bpcwip[(df_bpcwip["PP_NAME"].isin(temp_second_df_PP_NAME_list))]
            )
            # 前 n 個 temp_material_data_df_for_second_df
            temp_material_data_df_for_second_df = (
                temp_material_data_df_for_second_df.iloc[:n, :]
            )

            # drop by index
            dropped_index_list = temp_material_data_df_for_second_df.index.tolist()
            df_bpcwip = df_bpcwip.drop(dropped_index_list)
            print("-" * 50)
            pprint(temp_second_df_PP_NAME_list)
            pprint(dropped_index_list)
            print(temp_material_data_df_for_second_df)
            for j in dropped_index_list:
                output_data["LOC"].append(
                    temp_material_data_df_for_second_df.loc[j, "LOC"]
                )
                output_data["LOT"].append(
                    temp_material_data_df_for_second_df.loc[j, "LOT"]
                )
                output_data["EQ_NAME"].append(eq_name)
                output_data["PP_NAME"].append(
                    temp_material_data_df_for_second_df.loc[j, "PP_NAME"]
                )
                output_data["DEVICE"].append(device)

                nphh_list = df_second[
                    (df_second["DEVICE"] == device)
                    & (
                        df_second["PP_NAME"]
                        == temp_material_data_df_for_second_df.loc[j, "PP_NAME"]
                    )
                    & (df_second["MACHINE_NAME"] == eq_name)
                ]["NPPH"].values.tolist()
                output_data["NPPH"].append(
                    nphh_list[0] if len(nphh_list) > 0 else np.nan
                )
                output_data["Ranking"].append(count)
                count += 1

    df = pd.DataFrame(output_data)
    df.to_excel("output.xlsx", index=False)


if __name__ == "__main__":
    start_time = time.time()
    # step1()
    step2()
    # step3()
    print("--- %s seconds ---" % (time.time() - start_time))