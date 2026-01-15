### Alphanumeric Validator

> Purpose: Validates if a target string contains only Alphanumeric characters (0-9, A-Z). Includes toggleable options for spaces and empty strings.
> 
> Source: ERPM Dashboard (DB)
> 
> Context: Used in the Product Database Entry form to prevent SQL Injection risks or invalid SKU formats before they reach the VBA layer.

```excel-formula
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

#### Technique:

Uses `SEQUENCE` + `CODE` + `SUMPRODUCT` array logic to check every character byte-by-byte. 

#### Attribution & Logic:

- Original Concept: Adapted from a common Excel forum technique (SUM/FIND against a helper string).
- Refactor: I rewrote the logic completely to use ASCII Byte analysis (CODE/SEQUENCE).
- Improvements: Removed volatile functions (INDIRECT), eliminated external cell dependencies, and encapsulated it in a reusable LET/LAMBDA structure for performance.

### ASCII Validator

> Purpose: Validates that a string contains only standard ASCII characters (0-127).
> 
> Source: ERPM Dashboard (DB)
> 
> Context: Ensures data compatibility with barcode printers 

```excel-formula
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

# 


