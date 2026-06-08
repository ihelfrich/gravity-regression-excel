# Linear regression in Excel, one variable at a time

A working lesson built on 200 real country-pair trade flows. You will estimate three
regressions by hand in Excel, read every number they produce, and learn what each one
means. Open `Gravity_Regression_Workbook.xlsx` alongside this page and work the two
together. Nothing here is simulated; the data comes from BACI 2022 and the CEPII Gravity
database.

## The puzzle

Some pairs of countries trade enormous amounts. Germany and France move tens of billions
of dollars of goods between them. Other pairs barely trade at all. What decides the
difference?

There is an old and very durable answer, borrowed from physics, called the gravity model.
Two bodies attract in proportion to their masses and in inverse proportion to the distance
between them. Swap mass for economic size and the analogy holds up startlingly well for
trade: big economies trade more, and distant economies trade less. Regression lets us put
exact numbers on "more" and "less," and lets us check whether the story survives contact
with data.

We build the model in three passes. Each pass adds one idea.

## Pass 1: one predictor, and the limits of one predictor

Start with distance alone. Both trade and distance are in natural logs, so the relationship
we fit is

```
ln_trade = a + b · ln_dist + error
```

The slope `b` is an elasticity. If `b` is −0.8, then a pair of countries twice as far apart
(a 100 percent increase in distance) trades about 80 percent less.

In Excel you do not need a regression command for this. Three built-in functions do the
whole job:

```
slope      =SLOPE(ln_trade, ln_dist)
intercept  =INTERCEPT(ln_trade, ln_dist)
R-squared  =RSQ(ln_trade, ln_dist)
```

On our 200 pairs the slope comes out near −0.81. The sign is right and the effect is real:
the t-statistic is about −3.7, well past the rough threshold of 2 that we use for "unlikely
to be an accident." Distance matters.

And yet R-squared is only about 0.06. That number is the share of the variation in trade
that the line accounts for. Six percent is almost nothing. Distance is a genuine force, but
on its own it leaves the puzzle almost entirely unsolved. This is the most useful lesson in
the whole workbook: a coefficient can be real, statistically sharp, and still explain very
little by itself. A single variable is rarely the story.

One more thing worth doing once, on the workbook's L1 tab, is to confirm where the slope
comes from. `=COVARIANCE.P(ln_dist, ln_trade)/VAR.P(ln_dist)` returns the same number as
`SLOPE`. The slope is just the covariance of x and y divided by the variance of x. The
function is a convenience, not a black box.

## Pass 2: many predictors, and the word "holding fixed"

Now add the two economic sizes. The gravity equation proper is

```
ln_trade = a + b1 · ln_dist + b2 · ln_gdp_i + b3 · ln_gdp_j + error
```

where `i` is the exporter and `j` is the importer. With more than one predictor, `SLOPE`
will not serve. The tool is `LINEST`, which returns a whole block of output at once:

```
=LINEST(ln_trade, X_range, TRUE, TRUE)
```

The third argument `TRUE` tells Excel to fit an intercept. The fourth `TRUE` asks for the
full statistics, not just the coefficients. The result is a five-row block. There is one
trap, and it catches everyone the first time: **LINEST returns the coefficients in reverse
order**, last predictor first, with the intercept in the final column. The workbook untangles
this for you with `INDEX(LINEST(...), row, col)`, which plucks a single number out of the
block so you can label it properly.

Here is what the three coefficients say on our data:

| Term | Coefficient | Reading |
|---|---:|---|
| ln_dist | −0.83 | trade falls with distance, about as before |
| ln_gdp_i | 1.05 | exporter size: nearly one-for-one |
| ln_gdp_j | 0.89 | importer size: close behind |

Two things deserve attention. First, the two GDP elasticities sit close to 1. That is not a
coincidence we engineered; it is what the gravity model predicts, and it falls straight out
of the data. A 1 percent larger economy trades very roughly 1 percent more. Second, the
distance coefficient barely moved when we added GDP (−0.81 became −0.83). That tells us
distance was not secretly standing in for size. The two forces are largely separate.

The payoff is in R-squared. It climbs from about 0.06 to about 0.35. Adding economic size
took us from explaining almost none of the variation in trade to explaining a third of it.
Size, not distance, was the missing piece.

A coefficient in a multiple regression is a partial effect: the effect of one variable
holding the others fixed. The distance elasticity of −0.83 is what distance does to trade
between two economies of given sizes. That phrase, holding the others fixed, is the entire
reason multiple regression exists.

## Pass 3: yes-or-no variables

Not everything that shapes trade is a quantity. Some things are simply true or false. Do the
two countries share a border? Do they speak the same language? Is there a trade agreement
between them? Each of these is a dummy variable: 1 when the condition holds, 0 when it does
not. You add them to the regression exactly like any other column.

A dummy's coefficient is the gap between the two groups, holding everything else fixed.
Because the outcome is in logs, that gap reads as a percentage. A coefficient `b` means
trade is higher by roughly `b × 100` percent. For larger coefficients the exact figure is
`(e^b − 1) × 100` percent, which the workbook computes with `=EXP(b)-1`.

On the full model:

| Dummy | Coefficient | Exact effect on trade |
|---|---:|---:|
| agree_fta | 1.13 | about +208 percent |
| common_language | 0.55 | about +73 percent |
| contiguity | −1.81 | negative, see below |

A trade agreement goes with trade more than tripling. A shared language adds most of a
doubling. These are large and they make sense.

The border coefficient does not make sense, and that is the point of including it. A shared
border comes out negative here, as if neighbors trade less, which contradicts everything we
know. The honest reading is about the sample, not the world. With only 200 pairs, the
handful of neighbors in this set are also short-distance, high-GDP pairs, and once distance
and size are controlled the leftover border term is estimated off very little independent
variation. The lesson is not "borders reduce trade." The lesson is that a surprising sign on
a small sample is a prompt to get more data and look harder, not a finding to report. Real
data argues back, and a good analyst lets it.

## Reading the diagnostics

Three numbers tell you whether to trust a coefficient and whether to trust the model.

The **t-statistic** is the coefficient divided by its standard error. As a rule of thumb, a
t bigger than about 2 in size means the coefficient is unlikely to be zero by chance. The
workbook also reports the **p-value** with `=T.DIST.2T(ABS(t), df)`, which turns the t into
the probability of seeing an effect this large if the true effect were zero. Small p, around
0.05 or below, is the usual bar.

**R-squared** is the share of variation explained, from 0 to 1. It only ever rises when you
add variables, so for comparing models with different numbers of predictors use **adjusted
R-squared**, which charges a penalty for each variable and can fall if the variable was not
pulling its weight.

The **F-statistic** asks a single yes-or-no question about the whole regression: do the
predictors jointly explain anything at all? A large F with a small p-value says yes.

## Three ways to run a regression in Excel

You now have all three, and they agree to the last decimal because they are the same
calculation underneath.

1. **Trendline.** Make a scatter plot, right-click a point, add a linear trendline, and tick
   "Display equation" and "Display R-squared." Fast, visual, one predictor only. This is on
   the L1 tab.
2. **Functions.** `SLOPE`, `INTERCEPT`, `RSQ` for one predictor; `LINEST` for many. These are
   live: change a number in the data and every result updates. This is the skill worth
   owning.
3. **Analysis ToolPak.** Data, Data Analysis, Regression, point at the Y and X ranges, click
   OK. Menu-driven and familiar, but it produces a static snapshot that does not update when
   the data changes. Good for a one-off; weaker as a habit. The Start Here tab explains how
   to switch it on.

## What to take away

Build models one variable at a time and watch what each addition does. A real but lonely
predictor explained 6 percent of trade; adding economic size took it to 35 percent. Read the
coefficients as elasticities when the variables are logged and as percentage gaps when they
are dummies. Keep one eye on the diagnostics so you know which numbers to lean on. And when a
sign comes out backwards, treat it as a question, not an answer.

The exercises tab puts all of this in your hands. Work it before you read the answer key.
