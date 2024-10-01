# -*- coding: utf-8 -*-
"""
Created on Tue Oct  1 19:52:06 2024

@author: fisch
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
from reportlab.platypus import Spacer

sns.set_style("darkgrid")

# Lade die Excel-Datei
file_path = 'DXS_Closed_Positions_Report.xlsx'  # Ersetze dies mit dem Pfad zu deiner Datei
df = pd.read_excel(file_path)
df = df.dropna(subset=['Market'])
datetime_columns = ["Proposed at", "Opened at", "Closed at"]
df[datetime_columns] = df[datetime_columns].apply(pd.to_datetime)

# Calculating other useful columns
df['Cumulative Net P/L (USD)'] = df['Net P/L (USD)'][::-1].cumsum()[::-1]
df['Cumulative Net P/L (USD)'] = df['Cumulative Net P/L (USD)'].ffill()
df['Cumulative Gross P/L (USD)'] = df['Gross P/L (USD)'][::-1].cumsum()[::-1]
df['Cumulative Gross P/L (USD)'] = df['Cumulative Gross P/L (USD)'].ffill()
df["Duration"] = df["Closed at"] - df["Opened at"]
df["Duration in days"] = df["Duration"].dt.days
df["Net P/L (BSV)"] = df["Net P/L (USD)"] / df["Close BSV price"]
df["Gross P/L (BSV)"] = df["Gross P/L (USD)"] / df["Close BSV price"]
df['Cumulative Net P/L (BSV)'] = df['Net P/L (BSV)'][::-1].cumsum()[::-1]
df['Cumulative Gross P/L (BSV)'] = df['Gross P/L (BSV)'][::-1].cumsum()[::-1]
df["Margin (USD)"] = df["Margin (BSV)"] * df["Open BSV price"]

# Total number of trades and average trade duration in days
n_total = len(df["Net P/L (USD)"])
average_trade_duration = df["Duration"].mean().days

# Calculating win rate
PnL_positive = df[df["Net P/L (USD)"] > 0]
n_positive = len(PnL_positive["Net P/L (USD)"])
win_rate = n_positive / n_total
win_rate_perc = win_rate * 100

# Calculating loss rate
PnL_negative = df[df["Net P/L (USD)"] < 0]
n_negative = len(PnL_negative["Net P/L (USD)"])
loss_rate = n_negative / n_total
loss_rate_perc = loss_rate * 100

# Calculating the best and worst performing instruments
total_PnL_market = df.groupby('Market')['Net P/L (USD)'].sum().reset_index()
total_PnL_market = total_PnL_market.sort_values(by="Net P/L (USD)", ascending=False)
top_10 = total_PnL_market.head(10)
worst_10 = total_PnL_market.tail(10)

# Calculating average net returns on margin
average_net_return_on_margin = (df["Net P/L (BSV)"].mean() / df["Margin (BSV)"].mean()) * 100
average_daily_net_return_on_margin = ((1 + average_net_return_on_margin/100)**(1 / average_trade_duration) - 1)*100
average_monthly_net_return_on_margin = ((1 + average_daily_net_return_on_margin/100)**30 - 1) * 100
average_yearly_net_return_on_margin = ((1 + average_daily_net_return_on_margin/100)**365 - 1) * 100

# Calculating other statistics
mean_net_profit = PnL_positive["Net P/L (USD)"].mean() 
mean_net_loss = PnL_negative["Net P/L (USD)"].mean() 
mean_margin_usd = df["Margin (USD)"].mean()
mean_margin_bsv = df["Margin (BSV)"].mean()
average_trade_duration = df["Duration"].mean()
median_trade_duration = df["Duration"].median()
net_pnl_bsv = df['Cumulative Net P/L (BSV)'].head(1)
net_pnl_bsv = float(net_pnl_bsv)
gross_pnl_bsv = df['Cumulative Gross P/L (BSV)'].head(1)
gross_pnl_bsv = float(gross_pnl_bsv)
average_net_pnl_per_trade = df["Net P/L (USD)"].mean()

# Plot for Cumulative Net P/L
fig, ax1 = plt.subplots(figsize=(6,4))
# Plotting Cumulative Net P/L (USD)
sns.lineplot(x=df["Closed at"], y=df["Cumulative Net P/L (USD)"], ax=ax1, color="#2A52BE", linewidth=2)
# Creating a second y-axis for Cumulative Net P/L (BSV)
ax2 = ax1.twinx()
sns.lineplot(x=df["Closed at"], y=df["Cumulative Net P/L (BSV)"], ax=ax2, color="green", linewidth=2)
# Formatting the first axis
ax1.set_xlabel("Date")
ax1.set_ylabel("Net P/L in USD", color="darkblue")
ax1.tick_params(axis='y', labelcolor='darkblue')
ax1.tick_params(axis='x', rotation=45)
ax1.set_title("Cumulative Net P/L", fontweight='bold')
# Formatting the second axis
ax2.set_ylabel("Net P/L in BSV", color="darkgreen")
ax2.tick_params(axis='y', labelcolor='darkgreen')
# Adding legends
lines_1, labels_1 = ax1.get_legend_handles_labels()
lines_2, labels_2 = ax2.get_legend_handles_labels()
ax1.legend(lines_1 + lines_2, labels_1 + labels_2)
plt.tight_layout()  # Adjust layout for better fit
  # Adjust rect to leave space for the title
# Save the plot as an image
plt.savefig("net_pnl.png",  bbox_inches='tight')  # Save to a file
plt.close()

# Histplot for trade distribution per duration
plt.figure(figsize=(6, 4))  # Set the figure size to 6 inches wide and 3 inches tall
sns.histplot(df["Duration in days"], bins=100, kde=True, palette = "Blues")
plt.title("Distribution of Trades per Duration", fontweight='bold')
  # Adjust rect to leave space for the title
plt.savefig("histplot_duration.png", bbox_inches='tight')  # Save to a file
plt.close()

# Plotting the best and worst performing instruments
fig, (ax1,ax2) = plt.subplots(nrows=1, ncols=2, figsize=(6,3.5))
# Erstelle einen Bar Plot
sns.barplot(data=top_10, x='Market', y='Net P/L (USD)', ax=ax1,  color="#2CA02C", alpha=0.8)
ax1.set_xlabel("Best performing Instruments", fontweight="bold")
ax1.set_ylabel("Total Net P/L (USD)")
ax1.tick_params(axis='x', rotation=90)
# Erstelle einen Bar Plot
sns.barplot(data=worst_10, x='Market', y='Net P/L (USD)', ax=ax2,  color="#D62728", alpha=0.8)
ax2.set_xlabel("Worst performing Instruments", fontweight="bold")
ax2.set_ylabel("Total Net P/L (USD)")
ax2.tick_params(axis='x', rotation=90)
plt.tight_layout()

plt.savefig("market_performance.png", bbox_inches='tight')  # Save to a file
plt.close()

# Scatterplot for trades per duration
# Set x-axis to log scale
plt.figure(figsize=(6, 4))
sns.scatterplot(data=df, x="Duration in days", y = "Net P/L (USD)", color="#2A52BE", s=30, alpha=0.5)
plt.grid(True, which='both', linestyle='--', linewidth=0.5)
plt.title('Net P/L per Trade and Duration', fontweight='bold')
plt.savefig("scatter_duration.png",  bbox_inches='tight')
plt.close()

# Scatterplot for trades per duration in log-scale
# Set up the figure
plt.figure(figsize=(6, 4))
# Create a scatterplot
sns.scatterplot(data=df, x="Duration in days", y="Net P/L (USD)", color="#2A52BE", s=30, alpha=0.5)
# Set the axes to logarithmic scale
plt.xscale('log')  # Log scale for x-axis
# Add titles and labels
plt.title('Net P/L per Trade and Duration in Log Scale', fontweight='bold')
plt.xlabel('Duration in Days (Log Scale)')
plt.ylabel('Net P/L (USD)')
# Show the grid for better readability
plt.grid(True, which='both', linestyle='--', linewidth=0.5)
plt.savefig("scatter_duration_log.png", bbox_inches='tight')
plt.close()

# Create a dictionary of the statistics in the specified order for the table in the pdf
stats = {
    "Total number of trades": n_total,
    "Net P/L (USD)": round(df["Net P/L (USD)"].sum(), 2),
    "Net P/L (BSV)": round(net_pnl_bsv, 3),
    "Holding fees (USD)": round(df["Holding fee (USD)"].sum(), 2),
    "Gross P/L (USD)": round(df["Gross P/L (USD)"].sum(), 2),
    "Gross P/L (BSV)": round(gross_pnl_bsv, 3),
    "Average net P/L per trade (USD)": round(average_net_pnl_per_trade, 2),
    "Net win rate (%)":  round(win_rate_perc, 2),
    "Average net profit (USD)": round(mean_net_profit, 2),
    "Net loss rate (%)": round(loss_rate_perc, 2),
    "Average net loss (USD)": round(mean_net_loss, 2),
    "Average margin per trade (USD)": round(mean_margin_usd, 2),
    "Average margin per trade (BSV)": round(mean_margin_bsv, 4),
    "Average return on deployed margin per trade (%)": round(average_net_return_on_margin, 2),
    "Average daily net return on margin (%)": round(average_daily_net_return_on_margin, 4),
    "Average monthly net return on margin (%)": round(average_monthly_net_return_on_margin, 2),
    "Average yearly net return on margin (%)": round(average_yearly_net_return_on_margin, 2),
    "Average trade duration (days)": round(average_trade_duration.days, 4),   
}

# Prepare data for the PDF table
table_data = []
for key, value in stats.items():
    table_data.append([key, value])

# Create a PDF file
pdf_file = "Trade_Analysis_Report.pdf"
doc = SimpleDocTemplate(pdf_file, pagesize=A4)
styles = getSampleStyleSheet()

# Add a title with the current date
title = Paragraph(f"Trade Analysis Report ({datetime.now().strftime('%Y-%m-%d')})", styles['Title'])

# Define a spacer 
spacer = Spacer(width=0, height=10)

# Create a table from the table_data
table = Table(table_data)
table.setStyle(TableStyle([
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('BACKGROUND', (0, 0), (-1, -1), colors.white),
    ('GRID', (0, 0), (-1, -1), 1, colors.gray),
    ('FONTSIZE', (0, 0), (-1, -1), 8),  # Set font size for all cells to 10
]))

# Define the images for the report
net_pnl_img = Image("net_pnl.png")
histplot= Image("histplot_duration.png")
marketplot = Image("market_performance.png")
scatterdur = Image("scatter_duration.png")
scatterdurlog = Image("scatter_duration_log.png")

# Build the PDF
elements = [title, spacer, table, spacer, spacer, net_pnl_img, histplot, spacer, spacer, marketplot, scatterdur, spacer, spacer, scatterdurlog]
doc.build(elements)

print(f"PDF report generated: {pdf_file}")