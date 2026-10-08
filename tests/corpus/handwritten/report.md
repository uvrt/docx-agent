# Quarterly report

The third quarter was *good*, and some of it was **very good**: revenue grew
**14%**, costs fell by *3 %* and ~~nothing~~ little went wrong.[^growth]

## Results by region

| Region | Q2 | Q3 | Note |
| :--- | ---: | ---: | :---: |
| North | 3.9 | 4.1 | *steady* |
| South | 2.2 | 2.8 | **up** |
| East |  | 1.0 | new |

See [the appendix](#appendix) and the [company site](https://example.com/about "About us").

### What changed

1. Pricing moved to a yearly plan.
2. Two teams merged:
   - sales and marketing,
   - support and success.
3. The `billing-v2` service went live.

> We will keep the same course next year.
> Nothing in the plan changes.

#### Method

The figures are unaudited.\
They come from the ledger on 1 October.

```
SELECT region, sum(revenue)
  FROM ledger
 GROUP BY region;
```

---

##### Appendix

###### Sources

The ledger and the CRM export.[^sources]

[^growth]: Measured against the same quarter last year.
[^sources]: Both as of 1 October 2026.
