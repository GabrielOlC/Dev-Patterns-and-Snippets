### Dynamic 2D Lookup

> Purpose: Perform a 2D lookup (row & column intersection) by dynamically referencing table headers and totals.
> 
> Source: Calculadora de Precificação Sacos para Lixo (CPSL)
> 
> Contex: Used to retrieve specific material specifications and cost from 
>            a single table interface.

```excel-formula
/* -------------------------------------------------------------------------
   Requirements: 
     - Update all references to match the target table
     - Implement INDEX-based row/column indices
------------------------------------------------------------------------- */

=INDEX(
  tbSDimentionBags[#All],
  XLOOKUP([@Litragem], tbSDimentionBags[Litragem], tbSDimentionBags[Sup_Row]),
  XLOOKUP([@Material], tbSDimentionBags[#Headers], tbSDimentionBags[#Totals])
)
   → Returns the intersecting value for the given row (Litragem ) and column (Material) 
```

#### Helper formulas for row/column indices

* Row Index equation. Place this in a column:
  
  ```excel-formula
  =ROW([@Litragem]) - MIN(ROW([Litragem])) + 2
  ```
  
  → Generates sequential row numbers within the table
  
  
- Column Index equation. Place this in the total's row:
  
  ```excel-formula
  =COLUMN() - COLUMN(tbSDimentionBags[[#Headers],[Litragem]]) + 1
  ```
  
  → Generates sequential column numbers relative to the "Litragem" header
  
  

#### Future Improvements:

- [ ] Eliminate the need for the Index by `xmatch` the whole table for the headers values, and return its relative position.


