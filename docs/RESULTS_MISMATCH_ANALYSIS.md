# Results Mismatch Analysis

## Executive Summary

The RMSE values differ between `regression_ar1.py` (90.36) and the comprehensive script (173.85) for **three reasons**:

1. **Different aggregation methods** (90.36 vs 145.37)
2. **Different train/test splits** (145.37 vs 173.85)  
3. Combined effect: 90.36 vs 173.85 (92% difference)

## Detailed Breakdown

### Issue 1: Aggregation Method Difference

**regression_ar1.py approach:**
```
Step 1: Calculate RMSE for each of 30 employees separately
Step 2: Average the 30 per-employee RMSE values
Result: mean(RMSE_emp1, RMSE_emp2, ..., RMSE_emp30) = 90.36
```

**Comprehensive script approach:**
```
Step 1: Pool ALL predictions from all employees together (180 predictions)
Step 2: Calculate RMSE on the pooled predictions
Result: RMSE(all_predictions_pooled) = 173.85
```

**Why they differ:**  
RMSE is **not a linear function** due to the square root:

```
RMSE = sqrt(mean(errors²))
```

Averaging RMSEs ≠ RMSE of pooled data because `sqrt(mean(x))` ≠ `mean(sqrt(x))`

**Verification:**  
- Per-employee average: **90.36** ✓
- Pooled from individual script: **145.37** 
- Comprehensive script: **173.85**

Even when we pool regression_ar1.py predictions the same way as comprehensive, we get 145.37, not 173.85. This reveals a second difference...

---

### Issue 2: Different Train/Test Split Sizes

**regression_ar1.py:**
```
24 months total
├─ Training: months 1-18 (18 months)
└─ Testing:  months 19-24 (6 months)
```

**Comprehensive script:**
```
24 months total
├─ Training:   months 1-15 (15 months)
├─ Validation: months 16-18 (3 months) ← EXTRA SPLIT!
└─ Testing:    months 19-24 (6 months)
```

**Impact:**  
The comprehensive script uses **16.7% less training data** (15 vs 18 months).  
Less training data → worse model → higher RMSE

**This explains:**  
- regression_ar1.py pooled: 145.37 (trained on 18 months)
- Comprehensive pooled: 173.85 (trained on 15 months)
- Difference: **28.48 RMSE** or **19.6% worse**

---

## Complete Flow Chart

```
Actual Data (30 employees × 6 test predictions = 180 predictions)
                            │
            ┌───────────────┴────────────────┐
            │                                │
    regression_ar1.py              Comprehensive Script
    (Train on 18 months)           (Train on 15 months)
            │                                │
            ├─ Per-employee RMSE             ├─ Pool all predictions
            │  = 90.36 ✓                     │  
            │                                │
            └─ Pool predictions              └─ Calculate pooled RMSE
               = 145.37                         = 173.85
                    │                                │
                    └────────────┬───────────────────┘
                                 │
                          Difference: 28.48
                          (19.6% due to less training)
```

---

## Which Method Is Correct?

### Both Are Valid, But Measure Different Things

#### regression_ar1.py (Per-Employee Average = 90.36)
**What it measures:** "How well do we forecast the *average employee*?"
- ✓ Each employee weighted equally
- ✓ Good when all employees are equally important (HR/organizational view)
- ✗ Doesn't reflect total prediction error
- ✗ Dominated by employees with low variance

#### Comprehensive Script (Pooled = 173.85)
**What it measures:** "How well do we forecast *any salary prediction*?"
- ✓ Each prediction weighted equally
- ✓ Standard machine learning practice
- ✓ Reflects true prediction accuracy across all cases
- ✓ What sklearn metrics return by default
- ✗ Dominated by employees with high absolute errors

---

## Training Data Comparison

| Method | Train Months | Val Months | Test Months | Train Data % |
|--------|-------------|------------|-------------|--------------|
| **regression_ar1.py** | 18 | 0 | 6 | 75% |
| **Comprehensive** | 15 | 3 | 6 | 62.5% |

The comprehensive script holds out 3 months for validation (hyperparameter tuning), resulting in **12.5% less training data** available for the final model.

**Note:** Both scripts use the **same test set** (months 19-24), so they're evaluating on the same data. The difference is entirely due to training methodology.

---

## Recommendations

### For LaTeX Document

The document is **CORRECT** - it documents the comprehensive script results (173.85), which is the standard ML approach. The document should:

1. ✓ Keep current results (based on comprehensive script)
2. Add clarification that RMSE is calculated on **pooled predictions** (standard practice)
3. Mention that per-employee averaging would give different (lower) values
4. Explain why validation split reduces training data but improves methodology

### Which Script to Use?

**Use the comprehensive script** because:
1. ✓ Uses validation set for hyperparameter tuning (proper ML workflow)
2. ✓ Tests multiple evaluation methods (Train/Val/Test, Recursive, CV, Nested CV)
3. ✓ Uses standard pooled metrics (what sklearn does)
4. ✓ More rigorous and generalizable methodology

**regression_ar1.py** is useful for:
- Quick per-employee diagnostics
- Understanding individual employee forecast quality
- Debugging specific employee predictions

---

## Final Answer to "Why Are They Different?"

### Three Compounding Factors:

1. **Aggregation** (90.36 → 145.37): 
   - Per-employee average vs pooled calculation
   - Difference: **+60.8%**

2. **Training Size** (145.37 → 173.85):
   - 18 months vs 15 months training
   - Difference: **+19.6%**

3. **Combined Effect** (90.36 → 173.85):
   - Total difference: **+92.4%**

### Which Is "Correct"?

**Both are correct**, but the comprehensive script (173.85) is the **standard and recommended** approach because:
- Uses proper validation methodology
- Reports pooled metrics (ML standard)
- More conservative estimate of real-world performance
- Aligns with how sklearn and other ML libraries report metrics

The LaTeX document correctly documents the comprehensive approach and should be kept as-is, with minor clarifications about pooled vs per-employee metrics.

---

## Action Items

1. ✓ **Keep LaTeX document as-is** - it documents the correct methodology
2. Add brief note explaining pooled vs per-employee RMSE
3. Document that validation split reduces training data (methodologically correct)
4. Consider adding a section on "interpretation of RMSE values"
5. regression_ar1.py can be kept for diagnostic purposes but shouldn't be used for final results
