import pandas as pd


def Model_PPH(source_path):
    df = pd.read_excel(source_path)
    df = df.groupby(["MACHINE_NAME", "PP_NAME", "DEVICE"])["NPPH"].mean()
    c_df = pd.DataFrame(df)
    c_df.reset_index(inplace=True)
    c_df.to_excel(r"C:\Users\a0486121\Desktop\Second.xlsx", index=False)


if __name__ == "__main__":
    source_path = r"C:\Users\a0486121\Desktop\First.xlsx"
    Model_PPH(source_path)
