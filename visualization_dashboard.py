
import pandas as pd
import matplotlib.pyplot as plt

# Load the data
file_path = "Analytics_Quiz (3) (1) (1).xlsx"
df = pd.read_excel(file_path, sheet_name="Visualization", header=None, skiprows=2)
df.columns = ["Date", "Orders", "Defects", "Revenue", "Profit"]
df = df.dropna(how='all')
df["Date"] = pd.to_datetime(df["Date"], errors='coerce')
df[["Orders", "Defects", "Revenue", "Profit"]] = df[["Orders", "Defects", "Revenue", "Profit"]].apply(pd.to_numeric, errors='coerce')
df = df.dropna(subset=["Date"])
df["Profit Margin"] = df["Profit"] / df["Revenue"]

# Plot dashboard
fig, axes = plt.subplots(nrows=3, ncols=1, figsize=(12, 10), sharex=True)

axes[0].plot(df["Date"], df["Profit"], color="green")
axes[0].set_title("Profit Over Time")
axes[0].set_ylabel("Profit")
axes[0].grid(True)

axes[1].plot(df["Date"], df["Profit Margin"], color="blue")
axes[1].set_title("Profit Margin Over Time")
axes[1].set_ylabel("Margin")
axes[1].grid(True)

axes[2].plot(df["Date"], df["Defects"], color="red")
axes[2].set_title("Defects Over Time")
axes[2].set_ylabel("Defects")
axes[2].set_xlabel("Date")
axes[2].grid(True)

plt.tight_layout()
plt.savefig("matplotlib_dashboard.png")
plt.show()

# Export cleaned data to CSV for Tableau
df.to_csv("visualization_dashboard_data.csv", index=False)
