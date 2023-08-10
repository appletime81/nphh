import time
import numpy as np
import pandas as pd
import polars as pl
from copy import deepcopy
from pprint import pprint

pd.set_option("display.max_columns", None)


def step1():
    """
    # In WIP_ARRANGEMENT.xlsx, Red col can direct copy and paste to template_1,
    yellow col need to insert the formula in code like what we did in Excel
    """
    print("Processing step1...")
    df_wip_arrangement = pl.from_pandas(
        pd.read_excel(r"WIP_ARRANGEMENT.xlsx", sheet_name="DEV_DB")
    )

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

    print("Done step1!")


def step2():
    """
    In BPCWIP excel, need to insert the new col called "PP_name", logic as follwoing
    (Firstexcel[WIRE] == BPCWIP[WIREPN],
     Firstexcel[CAPILLARY] == BPCWIP[CAPILLARY],
     Firstexcel[Device] == BPCWIP[Device],
     If all same, bring the Firstexcel[PP_NAME] in this col)
    """
    print("Processing step2...")

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

    print("Done step2!")


def step3():
    print("Processing step3...")

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

    df_second = (
        pd.read_excel("Second.xlsx")
        .sort_values(by=["NPPH"], ascending=False)
        .reset_index(drop=True)
    )
    df_bpcwip = pd.read_excel("BPCWIP_2.xlsx").dropna(subset=["PP_NAME"])
    df_machine_setup = pd.read_excel("machine_setup.xlsx")

    for i in range(len(df_machine_setup)):
        machine_setup_pp_name = df_machine_setup.loc[i, "PP"]
        eq_name = df_machine_setup.loc[i, "EQ_NAME"]
        device = df_machine_setup.loc[i, "DEVICE"]

        temp_material_data_df = df_bpcwip[
            (df_bpcwip["PP_NAME"] == machine_setup_pp_name)
            & (df_bpcwip["STATUS"] == filter_condition["STATUS"])
            & (df_bpcwip["DEVICE"] == device)
        ]

        rows_to_add = temp_material_data_df.iloc[:3]

        if eq_name == "tain2c0016":
            print(rows_to_add)

        dropped_index_list = rows_to_add.index.tolist()
        df_bpcwip = df_bpcwip.drop(dropped_index_list)

        if len(rows_to_add) < 3:
            temp_second_df_PP_NAME_list = df_second[
                (df_second["MACHINE_NAME"] == eq_name)
            ]["PP_NAME"].values.tolist()
            temp_second_df_DEVICE_list = df_second[
                (df_second["MACHINE_NAME"] == eq_name)
            ]["DEVICE"].values.tolist()
            additional_rows = df_bpcwip[
                df_bpcwip["PP_NAME"].isin(temp_second_df_PP_NAME_list)
                & (df_bpcwip["DEVICE"].isin(temp_second_df_DEVICE_list))
            ].iloc[: 3 - len(rows_to_add)]

            # if eq_name == "tain2c0017":
            #     print("*" * 50)
            #     print(additional_rows)
            print("*" * 50)
            print(additional_rows)
            additional_rows_npph = []
            print("eq_name :", eq_name)
            for additional_row in additional_rows.iterrows():
                temp_pp_name = additional_row[1]["PP_NAME"]
                temp_device = additional_row[1]["DEVICE"]
                temp_npph = df_second[
                    (df_second["PP_NAME"] == temp_pp_name)
                    & (df_second["DEVICE"] == temp_device)
                    & (df_second["MACHINE_NAME"] == eq_name)
                ]["NPPH"].values.tolist()
                if not len(temp_npph):
                    temp_npph = df_second[
                        (df_second["PP_NAME"] == temp_pp_name)
                        & (df_second["DEVICE"] == temp_device)
                    ]["NPPH"].values.tolist()
                additional_rows_npph.append(temp_npph[0])
            additional_rows["NPPH"] = additional_rows_npph

            # sort by NPPH
            additional_rows = additional_rows.sort_values(by=["NPPH"], ascending=False)

            rows_to_add["NPPH"] = np.nan

            dropped_index_list = additional_rows.index.tolist()
            df_bpcwip = df_bpcwip.drop(dropped_index_list)
            rows_to_add = pd.concat([rows_to_add, additional_rows])

        count = 0
        for j, row in rows_to_add.iterrows():
            count += 1
            output_data["LOC"].append(row["LOC"])
            output_data["LOT"].append(row["LOT"])
            output_data["EQ_NAME"].append(eq_name)
            output_data["PP_NAME"].append(row["PP_NAME"])
            output_data["DEVICE"].append(row["DEVICE"])
            output_data["Ranking"].append(count)
            try:
                output_data["NPPH"].append(row["NPPH"])
            except KeyError:
                output_data["NPPH"].append("")

    df = pd.DataFrame(output_data)
    df.to_excel("output.xlsx", index=False)
    print("Done step3!")


if __name__ == "__main__":
    start_time = time.time()
    # step1()
    # step2()
    step3()
    print("--- %s seconds ---" % (time.time() - start_time))
