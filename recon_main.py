import pandas as pd
from itertools import combinations

def reconcile(df):
    reconciled_df = pd.DataFrame(columns=df.columns)
    to_drop = []

    for index, row in df.iterrows():
        if index in to_drop:
            continue

        amount = row['Amount']
        opposite_rows = df[(df['Amount'] == -amount) & (df.index != index) & (~df.index.isin(to_drop))]
        
        if not opposite_rows.empty:
            reconciled_df = pd.concat([reconciled_df, pd.DataFrame([row, opposite_rows.iloc[0]], columns=df.columns)])
            to_drop.extend(opposite_rows.index)
            to_drop.append(index)

    return reconciled_df, df.drop(to_drop)

def find_combinations(target, remaining_df):
    print(f"remaining_df.index size is {len(remaining_df.index)}")
    # TODO how to further improve the efficincey to handle one-to-many > 3
    for r in range(1, 3): # can only support one-to-many up to 3 now
        print(f"transaction combination size is {r}")
        for combo_indices in combinations(remaining_df.index, r):
            # Initialize sum for this combination
            combo_sum = 0
            # Iterate over each index in the combination
            for index in combo_indices:
                # Add the 'Amount' for this index to the sum
                combo_sum += remaining_df.loc[index, 'Amount']
            # combo_sum = remaining_df.loc[combo_indices, 'Amount'].sum()
            if combo_sum == target:
                print(f"Match found: combo_indices are {combo_indices}")
                return combo_indices
    return None

# Read the Excel file
df = pd.read_excel('BalanceSheetDetail_modified.xlsx')

# Reconcile until no further reconciliations can be made
reconciled_df = pd.DataFrame(columns=df.columns)
while True:
    reconciled_part, df = reconcile(df)
    if reconciled_part.empty:
        break
    reconciled_df = pd.concat([reconciled_df, reconciled_part])

# Attempt to reconcile rows by combining multiple rows
for index, row in df.iterrows():
    if index in reconciled_df.index:
        continue
    
    target = -row['Amount']
    remaining_df = df.drop(index)
    print(f"processing index {index} start")
    combo_indices = find_combinations(target, remaining_df)
    # print(f"combo_indices is {combo_indices}")
    if combo_indices:
        print(f"combo_indices is {combo_indices}")
        # Add the current row and rows in the combination to reconciled_df
        reconciled_df = pd.concat([reconciled_df, df.loc[[index, *combo_indices]]])
        # Drop the current row and rows in the combination
        df.drop([index, *combo_indices], inplace=True)

# Save the reconciled DataFrame to a new Excel file
reconciled_df.to_excel('reconciled_data.xlsx', index=False)

# Save the remaining unreconciled rows to another Excel file
df.to_excel('unreconciled_data.xlsx', index=False)
