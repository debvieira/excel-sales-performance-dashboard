# Sales Performance Dashboard — Excel

Excel project developed to consolidate monthly sales data and monitor revenue, targets, commissions, margins, and sales performance.

![Dashboard](imagens/dashboard.png)

## Project overview

The project simulates a company that receives a separate sales dataset every month. The monthly files are consolidated into a single analytical base and used to calculate commercial indicators.

The solution was developed in Excel Online without Power Query. Therefore, dynamic array formulas and native Excel features were used to create a semi-automated workflow.

## Key features

- Monthly sales consolidation using `VSTACK`
- Revenue, discount, cost, and margin calculations
- Sales target tracking
- Tier-based commission calculations
- Salesperson performance ranking
- Data quality validation
- Pivot tables and pivot charts
- Interactive filters and slicers
- Commercial performance dashboard
- Simulation of adding a new monthly dataset

## Workbook structure

| Worksheet | Purpose |
|---|---|
| `Vendas_Jan` to `Vendas_Abr` | Monthly sales datasets |
| `Base_Consolidada` | Consolidated data and calculated columns |
| `Cadastro_Vendedores` | Salesperson information |
| `Cadastro_Produtos` | Product prices and costs |
| `Cadastro_Metas` | Monthly targets |
| `Faixas_Comissao` | Commission rules |
| `Comissoes` | Monthly performance and commissions |
| `Ranking` | Salesperson ranking |
| `Validacao` | Data quality checks |
| `Analise_*` | Pivot tables and charts |
| `Dashboard` | Final commercial dashboard |

## Main calculation

```excel
=EMPILHARV(Vendas_Jan!A2:N43;Vendas_Fev!A2:N43;Vendas_Mar!A2:N43;Vendas_Abr!A2:N43)
```

This formula consolidates the four monthly sales datasets into a single dynamic array.

> The Portuguese version of Excel uses `EMPILHARV`, which is equivalent to `VSTACK` in English.

## Excel Online limitations

This project was developed using the free version of Excel Online, which has limitations compared to the desktop application.

Because Power Query was not available, the solution required:

- Manual addition of each monthly worksheet
- `VSTACK` to consolidate monthly datasets
- Manual extension of calculated columns
- Manual adjustment of fixed formula ranges
- Manual review of pivot table data sources
- Slicers kept in the analysis worksheets due to interface limitations

## How I would improve this project with Excel Desktop

With access to the desktop version of Excel, I would use Power Query to:

- Import all monthly files automatically from a folder
- Combine new sales files during refresh
- Standardize data types and column names
- Remove duplicates and blank records
- Identify invalid or missing values
- Merge sales with product and salesperson tables
- Create reusable transformation steps
- Generate a clean analytical dataset

After adding a new monthly file to the folder, the entire model could be updated using **Refresh All**, eliminating most manual adjustments and reducing the risk of errors.