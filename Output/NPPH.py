import pandas as pd
import time


def Model_PPH(source_path):
    df = pd.read_excel(source_path)
    df = df.groupby(["MACHINE_NAME", "PP_NAME", "DEVICE"])["NPPH"].mean()
    c_df = pd.DataFrame(df)
    c_df.reset_index(inplace=True)
    c_df.to_excel(r"Second.xlsx", index=False)


if __name__ == "__main__":
    start_time = time.time()
    source_path = "First.xlsx"
    Model_PPH(source_path)
    print("--- %s seconds ---" % (time.time() - start_time))
