import time
from copy import deepcopy

import pandas as pd
import polars as pl


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
    numbers = []
    nphh_list = []
    for i in range(len(df_machine_setup)):
        pp_name = df_machine_setup.loc[i, "PP"]
        numbers.append(
            len(
                df_bpcwip[
                    (df_bpcwip["PP_NAME"] == pp_name)
                    & (df_bpcwip["STATUS"] == "Non-RTD StandBy")
                ]
            )
        )
        nphh = df_template_1[
            df_template_1["DEVICE"] == df_machine_setup.loc[i, "DEVICE"]
        ]["NPPH"].values
        if len(nphh) == 0:
            nphh = None
        else:
            nphh = nphh[0]
        nphh_list.append(nphh)

    df_machine_setup["Numbers"] = numbers
    df_machine_setup["NPPH"] = nphh_list
    df_machine_setup.to_excel("machine_setup_2.xlsx", index=False)


if __name__ == "__main__":
    start_time = time.time()
    step3()
    print("--- %s seconds ---" % (time.time() - start_time))
