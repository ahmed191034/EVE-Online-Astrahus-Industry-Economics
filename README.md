# EVE Online — Astrahus Industry Economics Case Study

A manufacturing/operations case study built on a real EVE Online industry bill-of-materials export: how long does it take to build one Astrahus (a player-owned space station), how many additional factories would compress that to one week, what's the breakeven point, and should raw inputs be bought or manufactured.

## Files

| File | Description |
|---|---|
| `Space_station_challenge_Original.xlsx` | The dataset exactly as given — 4 open questions plus a 1,736-row manufacturing reference sheet. |
| `Space_station_challenge_Completed.xlsx` | The solved workbook: a full formula-driven build-time model plus the 4 answers, each pointing back to the exact sheet/formulas that produced it. |

## The dataset

Two sheets in the original file:

- **Space station challenge** — 4 open business questions, no data.
- **Reference sheet** — 1,736 rows across several unrelated column blocks: extraction/manufacturing time & throughput per production tier (Raw material → P1 → P2 → P3 → P4 → Space station structure), a 98-item name→tier lookup, factory prices per tier, the Astrahus blueprint (8 required structure components), and — critically — a 1,736-row **fully expanded bill of materials**: every raw ingredient needed to build each of the 8 blueprint components, flattened into a depth-first list (e.g. building "Structure Construction Parts" needs 450,000 Tritanium, 90,000 Pyerite, 3 Integrity Response Drones, which need Gel-Matrix Biopaste, which needs Biocells, which needs Biofuels and Carbon Compounds, and so on down to raw resources).

The workbook's "Input prices" block lists raw-material *category names* (Minerals, plus 15 PI resource categories) but was never populated with actual ISK-per-unit prices — only a link to evemarketer.com was left as a placeholder. That gap matters for questions 4 and 5 below, and is called out explicitly rather than papered over with invented numbers.

## Sheets added to the completed workbook

| New sheet | Purpose |
|---|---|
| `Tier Params` | The 7 production tiers with their time-per-cycle, output-per-cycle, and factory price, pulled directly from the Reference sheet. |
| `Item Tier Lookup` | All 98 named items mapped to their tier, with the exact source cell each mapping came from. |
| `BOM Cycle Calc` | All 1,733 raw bill-of-materials rows, each tagged with its Astrahus component, tier (via VLOOKUP), and per-row cycle/hour figures. |
| `Consolidated Item Demand` | The same data re-aggregated per (component, item) pair via SUMIFS — see note below on why this step is necessary — with cycles and hours computed from the aggregated quantity. |
| `Build Time & Factory Plan` | The two answer tables backing Q2 and Q3, entirely formula-driven. |

## Approach

**Reconstructing the build tree.** The 1,736-row Bill of Materials isn't one flat list — it's 8 separate depth-first traversals, one per Astrahus blueprint component (Structure Construction Parts, Hangar Array, Storage Bay, Repair Facility, Docking Bay, Medical Center, Office Center, Market Network), concatenated together. Each component's tree was located by finding where its own name reappears as a row (its root), giving 8 contiguous row ranges.

**Why quantities had to be aggregated before computing cycles.** The same ingredient (e.g. "Plasmoids") appears on multiple branches of a tree — once as an input to one sub-part, again as an input to a different sub-part. A first pass that computed `CEILING(quantity / output_per_cycle)` on each raw row independently overstated total cycles badly (46,679 hours instead of the correct 36,864), because splitting one item's real demand across several small rows and rounding each one up separately manufactures phantom extra cycles. The fix — the `Consolidated Item Demand` sheet — uses `SUMIFS` to total every item's real demand *within its component* first, and only then applies `CEILING`:

```
Quantity Required = SUMIFS('BOM Cycle Calc'!$D:$D, 'BOM Cycle Calc'!$B:$B, Component, 'BOM Cycle Calc'!$C:$C, ItemName)
Tier               = VLOOKUP(ItemName, 'Item Tier Lookup'!$A:$B, 2, FALSE)
Cycles             = CEILING(Quantity / Output_per_cycle, 1)
Hours              = Cycles * Time_per_cycle
```

This was caught by cross-checking the LibreOffice-recalculated total against an independent Python computation of the same tree — the two disagreed until the aggregation step was added, which is exactly the kind of error a "SUMIFS per raw row" formula design can hide silently.

## The answers

**Q1 — How long to build one Astrahus?**
Summing `Consolidated Item Demand` hours by tier (`Build Time & Factory Plan`, section 1, `SUMIF` by tier) gives the total processing time if one factory does every single job, one after another:

| Tier | Hours | Share |
|---|---|---|
| P1 | 18,519 | 50.2% |
| P2 | 9,168 | 24.9% |
| Raw material | 6,504 | 17.6% |
| P3 | 2,070 | 5.6% |
| P4 | 406 | 1.1% |
| Space station structure | 104 | 0.3% |
| Mineral | 93 | 0.3% |
| **Total** | **36,863.75 hrs** | **≈ 219.4 weeks (≈ 4.2 years)** |

**Q2 — How many additional factories, at what levels, to hit a 1-week build?**
Buying more factories lets multiple jobs at the same tier run in parallel, but tiers still depend on each other in sequence. The allocation that minimises total factory spend while bringing the sum of (tier hours ÷ factories at that tier) under 168 hours is a constrained optimisation, solved with a Lagrangian multiplier:

```
F[tier] = SQRT(Hours[tier] / Price[tier]) × (S / 168),  where S = Σ SQRT(Hours[tier] × Price[tier])
```

implemented as real formulas in `Build Time & Factory Plan`, section 2 (`SQRT`, `CEILING`, `MAX` — no hard-coded numbers). Result: **1,486 factories** across all 7 tiers (1,479 beyond a 1-per-tier baseline), costing **11,623** at the sheet's given factory prices, bringing total build time down to **164.2 hours** — just under the 1-week target:

| Tier | Factories | Resulting time | Cost |
|---|---|---|---|
| Mineral | 8 | 11.6 hrs | 800 |
| Raw material | 297 | 21.9 hrs | 1,485 |
| P1 | 646 | 28.7 hrs | 1,938 |
| P2 | 352 | 26.0 hrs | 1,760 |
| P3 | 133 | 15.6 hrs | 1,064 |
| P4 | 48 | 8.5 hrs | 576 |
| Space station structure | 2 | 52.0 hrs | 4,000 |

**Q3 — Breakeven point.**
The only cost figures actually present in the source data are the factory prices above and a "price per structure" of 2,000 for a completed Astrahus (Reference sheet, Structure pricing). Treating that 2,000 as net revenue per completed unit (material cost isn't separable — see the caveat below), breakeven on the 11,623 factory investment is **6 completed Astrahus structures** (`CEILING(11,623 / 2,000)`).

*Caveat, stated directly rather than glossed over:* this ignores raw-material cost, because the workbook's "Input prices" table was never populated with actual ISK-per-unit figures for the ~22 raw resources and minerals — only category names and a link to evemarketer.com were left. A materials-inclusive breakeven would need those live prices; fabricating plausible-looking numbers here would produce a confident-sounding but false answer, so the honest scope of this figure (factory investment only) is stated explicitly instead.

**Q4 — Buy or make input materials?**
Without per-unit market prices, this can't be answered as a $ margin comparison — but the workload breakdown above already gives a defensible operational answer. Raw material, P1 and P2 items together account for **92.7%** of all processing hours for the cheapest, highest-volume commodity tiers — exactly the materials EVE's market is normally deepest and cheapest for. Self-manufacturing them ties up enormous factory-time for low unit value, so the recommendation is: **buy** raw/P1/P2 inputs, **make** only P3/P4 and the final structure components (6.9% of total hours, where the value-add and the whole point of the exercise actually sit).

## Verification

The full workbook — all 1,733 BOM rows, the tier lookup, the aggregation, and both answer tables — was recalculated headlessly in LibreOffice Calc and the resulting cached values were checked against an independent Python re-implementation of the same tree-parsing and Lagrangian optimisation logic. Total build time (36,863.75 hrs), the per-tier breakdown, and the factory allocation (1,486 factories / 11,623 cost / 164.2 hrs) matched exactly between the two.
