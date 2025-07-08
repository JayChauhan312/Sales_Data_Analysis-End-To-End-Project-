### *Problem Statement*

''' XYZ Co's 2014-2018 sales data to identify key revenue and profit drivers across products, channels
 and regions, uncover seasonal trends and outliers, and align performance against budgets.
 Use these insights to optimize pricing, promotions and market expansion for sustainable growth
 and reduced concentration risk. I

#### *Objectives*

#Identify top performing products channels and regions driving revenue and profit'''

import pandas as pd
import numpy as np

import matplotlib.pyplot as plt
import seaborn as sns

'''Working with Excel Sheet
install Openpyxl'''

#excel_file = 'Regional Sales Dataset.xlsx'
import openpyxl
path = "D:/Youtube_project/Regional Sales Dataset.xlsx"
sheets = pd.read_excel(path,sheet_name = None)
print(sheets)

#Assign dataframes

df_sales = sheets['Sales Orders']
df_customers = sheets['Customers']
df_products = sheets['Products']
df_regions = sheets['Regions']
df_state_reg = sheets['State Regions']
df_budgets = sheets['2017 Budgets']

print(df_sales.shape)
print(df_sales.head(5))
print(df_sales.tail(5))

print("df_sales",df_sales.shape)
print("df_customers",df_customers.shape)
print("df_products",df_products.shape)
print("df_regions",df_regions.shape)
print("df_state_reg",df_state_reg.shape)
print("df_budgets",df_budgets.shape)

print(df_state_reg.head(5))






'''IMPORTANT'''
#Change heading to first row of state_reg

new_Header = df_state_reg.iloc[0]
df_state_reg.columns = new_Header
df_state_reg = df_state_reg[1:].reset_index(drop=True)
print(df_state_reg)
print(df_state_reg.head(5))

'''IMPORTANT'''





print("SALES : \n",df_sales.isnull().sum())
print("CUSTOMER : \n",df_customers.isnull().sum())
print("PRODUCT : \n",df_products.isnull().sum())
print("REGION : \n",df_regions.isnull().sum())
print("STATE_REGION : \n",df_state_reg.isnull().sum())
print("BUDGET : \n",df_budgets.isnull().sum())




# DATA CLEANING AND WRANGLING (MERGE Table AND build relation (Left join-Right join etc..))

#Merge with customers
df = df_sales.merge(
    df_customers,
    how = 'left',
    left_on = 'Customer Name Index',
    right_on = 'Customer Index'
)
print(df)

#Merge with products
df = df.merge(
    df_products,
    how = 'left',
    left_on = 'Product Description Index',
    right_on = 'Index'
)
print(df)



#Merge with regions Delivery Region Index
df = df.merge(
    df_regions,
    how = 'left',
    left_on = 'Delivery Region Index',
    right_on = 'id'
)
print(df)

#Merge with state_reg
df = df.merge(
    df_state_reg[["State Code","Region"]],
    how = 'left',
    left_on = 'state_code',
    right_on = 'State Code'
)
print(df)


#Merge with budgets
df = df.merge(
    df_budgets,
    how = 'left',
    on = 'Product Name',
)
print(df)

print(df.to_csv('file.csv'))




'''REMOVE COLUMNS State Code,id,Index,Customer Index'''
#Clean up redundant Column

col_drop = ['State Code','id','Index','Customer Index']
df = df.drop(columns = col_drop,errors= 'ignore')
print(df.head(5))



'''Convert all columns in to lower cases for consistency and easier access'''
df.columns = df.columns.str.lower()
print(df.columns.values)




''' Remove data which we thinks are not needed'''




col_need = [
    'ordernumber',        # unique order ID
    'orderdate',          # date when the order was placed
    'customer names',     # customer who placed the order
    'channel',            # sales channel (e.g., Wholesale, Distributor)
    'product name',       # product purchased
    'order quantity',     # number of units ordered
    'unit price',         # price per unit
    'line total',         # revenue for this line item (qty × unit_price)
    'total unit cost',    # company’s cost for this line item
    'state_code',         # two-letter state code
    'state',              # full state name
    'region',             # broader U.S. region (e.g., South, West)
    'latitude',           # latitude of delivery city
    'longitude',          # longitude of delivery city
    '2017 budgets'        # budget target for this product in 2017
]
df = df[col_need]
print(df.head(5))

df = df.rename(columns={
    'ordernumber'      : 'order_number',   # snake_case for consistency
    'orderdate'        : 'order_date',     # date of the order
    'customer names'   : 'customer_name',  # customer who placed it
    'product name'     : 'product_name',   # product sold
    'order quantity'   : 'quantity',       # units sold
    'unit price'       : 'unit_price',     # price per unit in USD
    'line total'       : 'revenue',        # revenue for the line item
    'total unit cost'  : 'cost',           # cost for the line item
    'state_code'       : 'state',          # two-letter state code
    'state'            : 'state_name',     # full state name
    'region'           : 'us_region',      # broader U.S. region
    'latitude'         : 'lat',            # latitude (float)
    'longitude'        : 'lon',            # longitude (float)
    '2017 budgets'     : 'budget'          # 2017 budget target (float)
})
print(df.head(1))



'''Blank out  budget for non-2017 orders'''
#df = df.loc[df['order_date'].dt.year != 2017,'budget'] = pd.NA

#line total is revenue
#print(df[['order_date','product_name','revenue','budget']].head(1))



# Blank out budgets for non-2017 orders
df.loc[df['order_date'].dt.year != 2017, 'budget'] = pd.NA

# Inspect
df[['order_date','product_name','revenue','budget']].head(10)
print(df[['order_date','product_name','revenue','budget']].head(10))
print(df[['order_date','product_name','revenue','budget']].tail(10))
print(df.info())




'''Filter the dataset to include only records from 2017'''

df_2017 = df[df['order_date'].dt.year == 2017]
df.isnull().sum()
print(df_2017)








'''Feature Engineering
    if we have col_1 and col_2 on bases of those cols we want to create col_3 it calls feature engineering
    whenever we want to add new feature it calls feature engineering'''

#df['total_cost'] = df['order_quantity'] * df['total unit cost']
#df['Profit'] = df['revenue'] - df['total_cost']
#df['Profit_margin_pct'] = df['Profit'] / df['revenue']*100
#print(df)


# 1. Calculate total cost for each line item
df['total_cost'] = df['quantity'] * df['cost']

# 2. Calculate profit as revenue minus total_cost
df['profit'] = df['revenue'] - df['total_cost']

# 3. Calculate profit margin as a percentage
df['profit_margin_pct'] = (df['profit'] / df['revenue']) * 100
print(df.head(5))








'''EDA'''
# Time series monthly chart and data
#method 1

df['order_month'] = df['order_date'].dt.to_period('m')

monthy_sales = df.groupby('order_month')['revenue'].sum()
plt.figure(figsize=(15,4))
monthy_sales.plot(marker='o',color = 'navy')
plt.title('Monthly Sales Revenue')
plt.xlabel('Month')
plt.ylabel('Revenue ($)')
plt.grid(True)
plt.tight_layout()
plt.show()

#method 2
'''
import pandas as pd
import matplotlib.pyplot as plt

# Convert order_date to datetime
df['order_date'] = pd.to_datetime(df['order_date'])

# Group by month and sum revenue
monthly_sales = df.groupby(df['order_date'].dt.to_period('M'))['revenue'].sum().reset_index()

# Convert period to string for plotting
monthly_sales['order_date'] = monthly_sales['order_date'].astype(str)

# Create the line chart
plt.figure(figsize=(12, 6))
plt.plot(monthly_sales['order_date'], monthly_sales['revenue'], marker='o')
plt.title('Monthly Sales Revenue')
plt.xlabel('Month')
plt.ylabel('Revenue ($)')
plt.xticks(rotation=45)
plt.grid(True)
plt.tight_layout()

# Show the plot
plt.show()
'''






'''
Sales trend by month
'''


import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Convert order_date to datetime
df['order_date'] = pd.to_datetime(df['order_date'])

# Extract month name for grouping
df['month'] = df['order_date'].dt.strftime('%B')

# Group by month and sum the revenue
monthly_sales = df.groupby('month')['revenue'].sum().reset_index()

# Ensure months are in calendar order
month_order = ['January', 'February', 'March', 'April', 'May', 'June',
               'July', 'August', 'September', 'October', 'November', 'December']
monthly_sales['month'] = pd.Categorical(monthly_sales['month'], categories=month_order, ordered=True)
monthly_sales = monthly_sales.sort_values('month')











# Create the line chart
plt.figure(figsize=(10, 6))
sns.lineplot(data=monthly_sales, x='month', y='revenue', marker='o')
plt.title('Overall Monthly Sales Trend (All Years Combined)')
plt.xlabel('Month')
plt.ylabel('Total Revenue')
plt.xticks(rotation=45)
plt.grid(True)
plt.tight_layout()
plt.show()










import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 1. Filter out any 2018 orders (NOT NEEDED 2018 DATA BECAUSE IT'S INCOMPLETE (Only have [jan-feb]))
df_N = df[df['order_date'].dt.year != 2018]

# Convert order_date to datetime
df_N['order_date'] = pd.to_datetime(df['order_date'])

# Extract month name for grouping
df_N['month'] = df['order_date'].dt.strftime('%B')

# Group by month and sum the revenue
monthly_sales = df_N.groupby('month')['revenue'].sum().reset_index()

# Ensure months are in calendar order
month_order = ['January', 'February', 'March', 'April', 'May', 'June',
               'July', 'August', 'September', 'October', 'November', 'December']
monthly_sales['month'] = pd.Categorical(monthly_sales['month'], categories=month_order, ordered=True)
monthly_sales = monthly_sales.sort_values('month')

# Create the line chart
plt.figure(figsize=(10, 6))
sns.lineplot(data=monthly_sales, x='month', y='revenue', marker='o')
plt.title('Overall Monthly Sales Trend (All Years Combined)')
plt.xlabel('Month')
plt.ylabel('Total Revenue')
plt.xticks(rotation=45)
plt.grid(True)
plt.tight_layout()
plt.show()








'''Find Top 10 products by revenue'''

top_prod = df.groupby('product_name')['revenue'].sum() / 1_000_000

# Select the top 10 products by revenue
top_prod = top_prod.nlargest(10)

# Set the figure size for clarity
plt.figure(figsize=(9, 4))

# Plot a horizontal bar chart: x-axis as revenue in millions, y-axis as product names
sns.barplot(
    x=top_prod.values,    # X-axis: revenue values in millions
    y=top_prod.index,     # Y-axis: product names
    palette='viridis'     # Color palette for bars
)

# Add title and axis labels
plt.title('Top 10 Products by Revenue (in Millions)')  # Main title
plt.xlabel('Total Revenue (in Millions)')              # X-axis label
plt.ylabel('Product Name')                             # Y-axis label

# Adjust layout to prevent overlapping elements
plt.tight_layout()

# Display the plot
plt.show()







'''Sales By Channel'''
chan_sales = df.groupby('channel')['revenue'].sum().sort_values(ascending=True)

plt.figure(figsize=(5,5))

plt.pie(
    chan_sales.values,
    labels=chan_sales.index,
    autopct='%1.1f%%'
)

plt.title('Total Sale By Channel')
plt.tight_layout()
plt.show()







'''
Average Order Value (AOV) Distribution'''

# Calculate the total revenue for each order to get the order value
aov = df.groupby('order_number')['revenue'].sum()

# Set the figure size for better visibility
plt.figure(figsize=(12, 4))

# Plot a histogram of order values
plt.hist(
    aov,  # Data: list of order values
    bins=50,  # Number of bins to group order values
    color='skyblue',  # Fill color of the bars
    edgecolor='black'  # Outline color of the bars
)

# Add title and axis labels for context
plt.title('Distribution of Average Order Value')
plt.xlabel('Order Value (USD)')
plt.ylabel('Number of Orders')

# Adjust layout to prevent clipping
plt.tight_layout()

# Show the plot
plt.show()








'''BOXPLOT Graph'''

# Set figure size for clarity
plt.figure(figsize=(12,4))

# Create a boxplot of unit_price by product_name
sns.boxplot(
    data=df,
    x='product_name',   # X-axis: product categories
    y='unit_price',      # Y-axis: unit price values
    color='g'            # Box color
)

# Add title and axis labels
plt.title('Unit Price Distribution per Product')  # Chart title
plt.xlabel('Product')                              # X-axis label
plt.ylabel('Unit Price (USD)')                     # Y-axis label

# Rotate x-axis labels for better readability
plt.xticks(rotation=45, ha='right')

# Adjust layout to prevent clipping of labels
plt.tight_layout()

# Display the plot
plt.show()








''''
Choropleth Map Chart'''

import plotly.express as px

# 1. Aggregate revenue by state (in millions)
state_sales = (
    df
    .groupby('state')['revenue']
    .sum()
    .reset_index()
)
state_sales['revenue_m'] = state_sales['revenue'] / 1e6  # convert to millions

# 2. Plotly choropleth
fig = px.choropleth(
    state_sales,
    locations='state',            # column with state codes
    locationmode='USA-states',    # tells Plotly these are US states
    color='revenue_m',
    scope='usa',
    labels={'revenue_m':'Total Sales (M USD)'},
    color_continuous_scale='Blues',
    hover_data={'revenue_m':':.2f'}  # show 2 decimals
)

# 3. Layout tuning
fig.update_layout(
    title_text='Total Sales by State',
    margin=dict(l=0, r=0, t=40, b=0),
    coloraxis_colorbar=dict(
        title='Sales (M USD)',
        ticksuffix='M'
    )
)

fig.show()









'''
2 Bar Charts '''

# 🔝 Calculate total revenue per customer and select top 10
top_rev = (
    df.groupby('customer_name')['revenue']
    .sum()  # Sum revenue for each customer
    .sort_values(ascending=False)  # Sort from highest to lowest
    .head(10)  # Keep top 10 customers
)

# 🔻 Calculate total revenue per customer and select bottom 10
bottom_rev = (
    df.groupby('customer_name')['revenue']
    .sum()  # Sum revenue for each customer
    .sort_values(ascending=True)  # Sort from lowest to highest
    .head(10)  # Keep bottom 10 customers
)

# Create a figure with two side-by-side subplots
fig, axes = plt.subplots(1, 2, figsize=(16, 5))

# Plot 1: Top 10 customers by revenue (converted to millions)
sns.barplot(
    x=top_rev.values / 1e6,  # X-axis: revenue in millions
    y=top_rev.index,  # Y-axis: customer names
    palette='Blues_r',  # Color palette (reversed blues)
    ax=axes[0]  # Draw on the left subplot
)
axes[0].set_title('Top 10 Customers by Revenue', fontsize=14)  # Title
axes[0].set_xlabel('Revenue (Million USD)', fontsize=12)  # X-axis label
axes[0].set_ylabel('Customer Name', fontsize=12)  # Y-axis label

# Plot 2: Bottom 10 customers by revenue (converted to millions)
sns.barplot(
    x=bottom_rev.values / 1e6,  # X-axis: revenue in millions
    y=bottom_rev.index,  # Y-axis: customer names
    palette='Reds',  # Color palette (reds)
    ax=axes[1]  # Draw on the right subplot
)
axes[1].set_title('Bottom 10 Customers by Revenue', fontsize=14)  # Title
axes[1].set_xlabel('Revenue (Million USD)', fontsize=12)  # X-axis label
axes[1].set_ylabel('Customer Name', fontsize=12)  # Y-axis label

# Adjust layout to prevent overlap and display both charts
plt.tight_layout()
plt.show()






'''
Bar Chart Bivariate'''

# 1️⃣ Compute average profit margin percentage for each channel
channel_margin = (
    df.groupby('channel')['profit_margin_pct']  # Group by sales channel
    .mean()  # Calculate mean profit margin %
    .sort_values(ascending=False)  # Sort channels from highest to lowest margin
)

# 2️⃣ Set the figure size for clarity
plt.figure(figsize=(6, 4))

# 3️⃣ Plot a bar chart of average profit margin by channel
ax = sns.barplot(
    x=channel_margin.index,  # X-axis: channel names
    y=channel_margin.values,  # Y-axis: average profit margin values
    palette='coolwarm'  # Color palette for bars
)

# 4️⃣ Add chart title and axis labels
plt.title('Average Profit Margin by Channel')  # Main title
plt.xlabel('Sales Channel')  # X-axis label
plt.ylabel('Avg Profit Margin (%)')  # Y-axis label

# 5️⃣ Annotate each bar with its exact margin percentage
for i, v in enumerate(channel_margin.values):
    ax.text(
        i,  # X position (bar index)
        v + 0.5,  # Y position (bar height + small offset)
        f"{v:.2f}%",  # Text label showing percentage with two decimals
        ha='center',  # Center-align the text horizontally
        fontweight='bold'  # Bold font for readability
    )

# 6️⃣ Adjust layout to prevent clipping and display the plot
plt.tight_layout()
plt.show()









'''Horizontal Bar Chart'''

# Aggregate total revenue and unique order count per state
state_rev = df.groupby('state_name').agg(
    revenue=('revenue', 'sum'),  # Sum up revenue per state
    orders=('order_number', 'nunique')  # Count unique orders per state
).sort_values('revenue', ascending=False).head(10)  # Keep top 10 by revenue

# Plot 1: Top 10 states by revenue (scaled to millions)
plt.figure(figsize=(15, 4))
sns.barplot(
    x=state_rev.index,  # X-axis: state names
    y=state_rev['revenue'] / 1e6,  # Y-axis: revenue in millions
    palette='coolwarm'  # Color palette
)
plt.title('Top 10 States by Revenue')  # Chart title
plt.xlabel('State')  # X-axis label
plt.ylabel('Total Revenue (Million USD)')  # Y-axis label
plt.tight_layout()  # Adjust layout
plt.show()  # Display the plot

# Plot 2: Top 10 states by number of orders
plt.figure(figsize=(15, 4))
sns.barplot(
    x=state_rev.index,  # X-axis: state names
    y=state_rev['orders'],  # Y-axis: order counts
    palette='coolwarm'  # Color palette
)
plt.title('Top 10 States by Number of Orders')  # Chart title
plt.xlabel('State')  # X-axis label
plt.ylabel('Order Count')  # Y-axis label
plt.tight_layout()  # Adjust layout
plt.show()  # Display the plot









'''
Customer segmentation - Multivariate '''

# Aggregate metrics per customer
cust_summary = df.groupby('customer_name').agg(
    total_revenue=('revenue', 'sum'),
    total_profit=('profit', 'sum'),
    avg_margin=('profit_margin_pct', 'mean'),
    orders=('order_number', 'nunique')
)

# Convert revenue to millions
cust_summary['total_revenue_m'] = cust_summary['total_revenue'] / 1e6

plt.figure(figsize=(7, 5))

# Bubble chart with revenue in millions
sns.scatterplot(
    data=cust_summary,
    x='total_revenue_m',        # <-- use revenue in millions
    y='avg_margin',
    size='orders',
    sizes=(20, 200),
    alpha=0.7
)

plt.title('Customer Segmentation: Revenue vs. Profit Margin')
plt.xlabel('Total Revenue (Million USD)')  # <-- updated label
plt.ylabel('Avg Profit Margin (%)')

plt.tight_layout()
plt.show()








'''
Corelation Heatmap'''

# List numeric columns to include in the correlation calculation
num_cols = ['quantity', 'unit_price', 'revenue', 'cost', 'profit']

# Calculate the correlation matrix for these numeric features
corr = df[num_cols].corr()

# Set the figure size for clarity
plt.figure(figsize=(6,4))

# Plot the heatmap with annotations and a viridis colormap
sns.heatmap(
    corr,           # Data: correlation matrix
    annot=True,     # Display the correlation coefficients on the heatmap
    fmt=".2f",      # Format numbers to two decimal places
    cmap='viridis'  # Color palette for the heatmap
)

# Add title for context
plt.title('Correlation Matrix')

# Adjust layout to prevent clipping
plt.tight_layout()

# Display the heatmap
plt.show()

