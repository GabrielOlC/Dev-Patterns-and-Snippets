### Advanced Pricing Algorithm (Fallback Logic)

> Purpose: Retrieves a recommended price based on multiple product attributes. Implements a "Fallback" mechanism: if the desired price tier (e.g., Tier 4) is missing, it automatically attempts to fetch the next available lower tier (Tier 3, then 2...).
> 
> Source: Calculadora de precificação sacos para lixo (CPSL)
> 
> Context: Solves the "Missing Price" issue in dynamic quoting. If a specific competitor price is missing, it defaults to standard pricing logic to ensure the user always sees a number, flagging it as an estimate       (returned as negative).

```excel-formula
/* -------------------------------------------------------------------------
   // Requirements:
      - Needs set conditional formating to treat negative values.
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

#### Future Improvements:

- [ ] Feature: Add a boolean parameter `vReturnNegative` to toggle whether fallback results are returned as negative numbers or flagged via a separate column.

- [ ] Optimization: Refactor the hardcoded `CHOOSE` columns into a dynamic array reference to handle table structural changes automatically.
  
  

#### Technique:

Uses `MAP` + `SEQUENCE` +` CHOOSE` to iterate backwards through price tiers.




