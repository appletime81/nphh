import pandas as pd


def Model_PPH(source_path):
    df = pd.read_excel(source_path)
    df = df.groupby(["MACHINE_NAME", "PP_NAME", "DEVICE", "WIRE"])["NPPH"].mean()
    c_df = pd.DataFrame(df)
    c_df.reset_index(inplace=True)
    c_df.to_excel(r"C:\Users\a0486121\Desktop\Second.xlsx", index=False)


def WIP_conversion(wip_excel):
    df = pd.read_excel(wip_excel)
    df = df.drop_duplicates(subset=["PP_NAME", "MACHINE_NAME", "DEVICE"], keep="last")
    df.to_excel(r"C:\Users\a0486121\Desktop\BPC_WIP_2.xlsx", index=False)


if __name__ == "__main__":
    source_path = r"C:\Users\a0486121\Desktop\Output\First.xlsx"
    # wip_excel = r'C:\Users\a0486121\Desktop\Output\First.xlsx'
    Model_PPH(source_path)
    # WIP_conversion(wip_excel)
