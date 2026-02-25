# Review: Transformer Remaining Capacity and Lockout Issues

## Likely root causes found

1. **Lockout margin applied in the wrong direction** in the original logic (`cap * (1 - LockoutMargin)`).
   - If users enter a positive margin (e.g., `3`), this lowered trip threshold to 97% instead of raising/offsetting it.
2. **Base load (`BL`) hardcoded to Station column**, even though workbook variants may place keys/values differently between `State` and `Station` columns.
3. **State normalization missing for verbose labels** (`RUNNING`, `STARTING`, etc.), causing lookup misses and resulting in zero loads.
4. **Lookup cache not refreshed** from the Refresh button, which could leave calculations stale after edits on `Load Analysis`.

## What was provided

A revised VBA module in `TransformerControl_Fixed.bas` that:
- Normalizes states to `PH/ST/RN/RD/BL`.
- Reads BL from either Station or State column (fallback).
- Applies lockout threshold as `cap * (1 + LockoutMargin)`.
- Clears lookup cache during refresh.
- Preserves your existing timing/transition behavior.

