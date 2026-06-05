# Task 1
📊 Excel – Data Exploration

📁 About the Dataset

I worked with a dataset containing **34 products**. Here is what each column in the dataset represents:

| Column | Description |
|--------|-------------|
| **Product ID** | Unique ID for each product (e.g., 28-JAN-US) |
| **Product Name** | Name of the product (e.g., Laptop, Sneakers) |
| **Brand Name** | Brand of the product (e.g., Dell, Nike) |
| **Price ($)** | Price of the product in USD |
| **Quantity** | Number of units |
| **Category** | Product category (Electronics, Fashion, Kitchen, Outdoor, Accessories) |

---

🔍 What I Did — Step by Step

1️⃣ Calculated the Total, Count, and Average Price

**Total price of all products:**
```
=SUM(D2:D35)  →  $10,100
```

**How many products are in the dataset:**
```
=COUNTA(B2:B35)  →  34 products
```

**Average price across all products:**
```
=AVERAGE(D2:D35)  →  ~$297.06
```

2️⃣ Found the Cheapest and Most Expensive Product

**Cheapest product:**
```
=MIN(D2:D35)  →  $30
```

**Most expensive product:**
```
=MAX(D2:D35)  →  $1,000
```

3️⃣ Created a Price Range Column Using IF

I created a new column called **Price Range** where I labeled each product based on its price. I used the IF function for this:
- If the price is **$500 or more** → I labeled it **High Price**
- If the price is **less than $500** → I labeled it **Standard Price**

```
=IF(D2>=500,"High Price","Standard Price")
```

4️⃣ Used SUMIF and COUNTIF for Specific Insights

**Total price of only Electronics products:**
```
=SUMIF(F2:F35,"Electronics",D2:D35)  →  $8,050
```

**How many products cost less than $100:**
```
=COUNTIF(D2:D35,"<100")  →  11 products
```

5️⃣ Extracted Information from the Product ID Column

On top of the original data, I created **4 new columns**:

| New Column       | What I did                                                                       |
|------------------|----------------------------------------------------------------------------------|
| **Price Range**  | Categorized each product as High Price or Standard Price using the `IF` function |
| **Day**          | Pulled the first 2 characters from the Product ID using the `LEFT` function      |
| **Country Code** | Pulled the last 2 characters from the Product ID using the `RIGHT` function      |
| **Month**        | Pulled the 4th to 6th characters from the Product ID using the `MID` function    |

Each ID follows the format `DD-MMM-CC` (e.g., `28-JAN-US`) which contains the day, month, and country code. 
So I used text functions to pull each part out into its own column:

Day — I took the first 2 characters:
```
=LEFT(A2,2)  →  e.g., "28"
```

Month — I took the 4th to 6th characters:
```
=MID(A2,4,3)  →  e.g., "JAN"
```

Country Code — I took the last 2 characters:
```
=RIGHT(A2,2)  →  e.g., "US"
```


📊 Final Results at a Glance

| Metric                  | Result   |
|-------------------------|----------|
| Total Price             | $10,100  |
| No. of Products         | 34       |
| Average Price           | ~$297.06 |
| Minimum Price           | $30      |
| Maximum Price           | $1,000   |
| Total Electronics Price | $8,050   |
| Products under $100     | 11       |

🛠️ Tools Used

-> Microsoft Excel — for all data exploration and formula work
-> GitHub — for uploading and sharing this project
