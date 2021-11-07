import argparse

import pandas as pd


def filter_workers(df):
    physics_workers = df[df["Area"].str.contains("Physics").fillna(False)]
    PIs_with_physics_workers = physics_workers["PI"].dropna().unique()
    workers_in_labs_with_physics = df[df.isin({"PI": PIs_with_physics_workers}).any(1)]
    # Combine area and PI data sets and deduplicate
    all_physics = pd.concat([physics_workers, workers_in_labs_with_physics])
    all_physics = all_physics[~all_physics.index.duplicated()]
    return all_physics


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("infile")
    parser.add_argument("outfile")
    args = parser.parse_args()
    df = pd.read_csv(args.infile)
    result = filter_workers(df)
    result.to_csv(args.outfile)
