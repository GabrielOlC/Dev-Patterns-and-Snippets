## Excel Snippets

***

## Lambda

> Requires that ones saves it named function (Formulas → Name Manager → New)

### Not Empty or Zero

```excel-formula
/* -------------------------------------------------------------------------
   Source: Calculadora de precificação sacos para lixo (CPSL)
   Name: lmdNotEmptyOrZero (Alias: lmdNEZ)
   Purpose: Helper function to validate that a cell contains meaningful data 
            (neither blank nor zero).
   Context: Used extensively in pricing logic to evaluate which logic persue
   ------------------------------------------------------------------------- */

=LAMBDA(lmCell, IF(AND(lmCell<>0, lmCell<>""), TRUE, FALSE))
```

## Normal

> Normal excel formulas

### Advanced Pricing Algorithm (Fallback Logic)

```excel-formula
/* -------------------------------------------------------------------------
   Source: Calculadora de precificação sacos para lixo (CPSL)
   Purpose: Retrieves a recommended price based on multiple product attributes. 
            Implements a "Fallback" mechanism: if the desired price tier 
            (e.g., Tier 4) is missing, it automatically attempts to fetch the 
            next available lower tier (Tier 3, then 2...).
   Context: Solves the "Missing Price" issue in dynamic quoting. If a specific 
            competitor price is missing, it defaults to standard pricing logic 
            to ensure the user always sees a number, flagging it as an estimate 
            (returned as negative).
   Technique: Uses MAP + SEQUENCE to iterate backwards through price tiers.
   Requirements: 
       - Needs set conditional formating to treat negative values.
       - set the references within the comments below.

   // Future Improvements:
   -[ ] Feature: Add a boolean parameter `vReturnNegative` to toggle whether 
        fallback results are returned as negative numbers or flagged via a separate column.
   -[ ] Optimization: Refactor the hardcoded `CHOOSE` columns into a dynamic 
        array reference to handle table structural changes automatically.
   ------------------------------------------------------------------------- */

=LET(
/* Needs change the conditions references and remove the comments {''} */
  vConditions,
         ($G8 {'Product name'} = tbPCCalculator[Litragem]) *
         (IF($B$5, $C$5 {'Thickness 1'}, $D$5 {'Thickness 2'}) = tbPCCalculator[Espessura]) *
         (tbPCCalculator[Und] = 100) *
         ($B$3 {'Material'} = tbPCCalculator[Material]) *
         (tbPCCalculator[Status] = "Ativo"),

  vCountValidConditions, SUMPRODUCT(vConditions),

/* Need define Pricing Tiers (Indices) - this is what tells which price range to use! */
  vOrigIndex, $B$2 {'price to use'} + 1, 

  vFallBackIndices, IF(vOrigIndex>=2, SEQUENCE(vOrigIndex-1, 1, vOrigIndex, -1), IF(vOrigIndex=1,{1},{0})),

/* Need define the Column's price names for the fallback*/
  vCandidateArray,
    MAP(vFallBackIndices,
      LAMBDA(lmIndex,
        IF(lmIndex=0, 0,
           SUMPRODUCT(vConditions, CHOOSE(lmIndex,
             tbPCCalculator[Preço de venda Recomendado],
             tbPCCalculator[PV Client Final],
             tbPCCalculator[PV Sem compromisso],
             tbPCCalculator[PV Venda em volume],
             tbPCCalculator[PV Abatendo concorrentes]
           ))
        )
      )
    ),

  vFirstCandidateEmpty, IF(INDEX(vCandidateArray,1)=0, TRUE, FALSE),
  vValidCandidates, FILTER(vCandidateArray, vCandidateArray<>0, 0),
  vResultCandidate, IF(AND(ROWS(vValidCandidates)=1, INDEX(vValidCandidates,1)=0), 0, INDEX(vValidCandidates, 1)),

  IF(vCountValidConditions >= 2, "Err. Unexpected duplicate",
    IF(vCountValidConditions = 0, "",
      IF(vFirstCandidateEmpty, vResultCandidate * -1, vResultCandidate)
    )
  )
)
```

### Alphanumeric Validator

```excel-formula
/* -------------------------------------------------------------------------
   Source: ERPM Dashboard (DB)
   Purpose: Validates if a target string contains only Alphanumeric characters 
            (0-9, A-Z). Includes toggleable options for spaces and empty strings.
   Context: Used in the Product Database Entry form to prevent SQL Injection 
            risks or invalid SKU formats before they reach the VBA layer.
   Technique: Uses SEQUENCE + CODE + SUMPRODUCT array logic to check every 
              character byte-by-byte.

   Attribution & Logic:
       - Original Concept: Adapted from a common Excel forum technique 
         (SUM/FIND against a helper string).
       - Refactor: I rewrote the logic completely to use ASCII Byte analysis 
         (CODE/SEQUENCE).
       - Improvements: Removed volatile functions (INDIRECT), eliminated external 
         cell dependencies, and encapsulated it in a reusable LET/LAMBDA structure 
         for performance.
   ------------------------------------------------------------------------- */

=LET(
  REM,n("This returns TRUE if target has only 0 to 9 and A to Z (upper or lowercase) characters. You may set to allow or not empty strings and Spaces"),
  vTarget, "As String or Range",
  vAllowEmptyTarget, FALSE,
  vAllowSpace, TRUE,

  lmdConRemoveSpace, LAMBDA(lmTxt, IF(vAllowSpace, SUBSTITUTE(lmTxt," ",""), lmTxt)),
  vTargetArray, CODE(MID(
                  lmdConRemoveSpace(vTarget), 
                  SEQUENCE(LEN(lmdConRemoveSpace(vTarget)))
                  , 1
                )),

  IF(vTarget <> "",
    IF(SUMPRODUCT(--( 
       (vTargetArray>=65)*(vTargetArray<=90) + 
       (vTargetArray>=97)*(vTargetArray<=122) + 
       (vTargetArray>=48)*(vTargetArray<=57) )) = LEN(lmdConRemoveSpace(vTarget)),
      TRUE, FALSE
    ),
    vAllowEmptyTarget
  )
)
```

### ASCII Validator

```excel-formula
/* -------------------------------------------------------------------------
   Source: ERPM Dashboard (DB)
   Purpose: Validates that a string contains only standard ASCII characters (0-127).
   Context: Ensures data compatibility with barcode printers 
   ------------------------------------------------------------------------- */

=LET(
  REM,n("This returns TRUE if the character is at ASCII"),
  vTarget, "As String or Range",
  vAllowEmptyTarget, FALSE,

  vTargetArray, CODE(MID(vTarget, SEQUENCE(LEN(vTarget)), 1)),

  IF(vTarget <> "",
    IF(MAX(vTargetArray) <= 127,
      TRUE, FALSE
    ),
    vAllowEmptyTarget
  )
)
```

### Dynamic 2D Lookup

```excel-formula
/* -------------------------------------------------------------------------
   Source: Calculadora de Precificação Sacos para Lixo (CPSL)
   Purpose: Perform a 2D lookup (row & column intersection) by dynamically 
            referencing table headers and totals.
   Contex: Used to retrieve specific material specifications and cost from 
           a single table interface.
   Requirements: 
     - Update all references to match the target table
     - Implement INDEX-based row/column indices

   // Future Improvements:
     -[ ] Eliminate the need for the Index by xlookup the whole table for the
          headers values, and return its relative position. (xmatch maybe?)
------------------------------------------------------------------------- */

/* Helper formulas for row/column indices */
Row Index, place this in a column:
=ROW([@Litragem]) - MIN(ROW([Litragem])) + 2
   → Generates sequential row numbers within the table

Column Index, place this in the total's row:
=COLUMN() - COLUMN(tbSDimentionBags[[#Headers],[Litragem]]) + 1
   → Generates sequential column numbers relative to the "Litragem" header

/* 2D lookup formula */
=INDEX(
  tbSDimentionBags[#All],
  XLOOKUP([@Litragem], tbSDimentionBags[Litragem], tbSDimentionBags[Sup_Row]),
  XLOOKUP([@Material], tbSDimentionBags[#Headers], tbSDimentionBags[#Totals])
)
   → Returns the intersecting value for the given row (Litragem ) and column (Material)


```
