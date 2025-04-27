# 📊 Excel Formulas for Finance and Accounting

A collection of core Excel formulas commonly used in financial analysis, accounting workflows, and business intelligence reporting.

---

## 🎯 Lookup and Search Formulas

### 1. INDEX MATCH
**Purpose:**  
Finds a value in a table based on a row and column lookup.  
More flexible and powerful than VLOOKUP.

**Example:**
```excel
=INDEX(B2:B10, MATCH(5000, A2:A10, 0))
```

Use Case:
Finding contract terms, customer pricing, or revenue figures without relying on column order. VLOOKUP is for quick lookups when table structure is simple and won’t change (which I rarely use now).



