### Not Empty or Zero

> Purpose: Helper function to validate that a cell contains meaningful data (neither blank nor zero).
> 
> Source: Calculadora de precificação sacos para lixo (CPSL)
> 
> Context: Used extensively in pricing logic to evaluate which logic persue

```excel-formula
/* Call: lmdNotEmptyOrZero (Alias: lmdNEZ)
   ---------------------------------------
   // Requirements
      - Requires that one saves it named function (Formulas → Name Manager → New)
   ------------------------------------------------------------------------- */

=LAMBDA(lmCell, IF(AND(lmCell<>0, lmCell<>""), TRUE, FALSE))
```


