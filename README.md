# profitability_calculator
This program takes the excel specification as input. Specification consists of articles and quantities.
It returns a table indicating discounts and the profitability. 
Input (excel sheet):
|-----------|---------------------------|----------|----------------|------------|
| article   | Denomination              | Quantity | Price per each | Price      |
| 831122466 | Usher, привод без тормоза | 1        | 8.397,74       | 8.397,74   | 
| 831122501 | БП USHER                  | 1        | 9.380,28       | 9.380,28   |
| 831123124 | Боковое ограждение Usher  | 22       | 640,23         | 14.085,06  |
| 831123125 | Бок огражд Usher 1-й сегм | 1        | 892,42         | 892,42     |
| 831123149 | Usher компл. зап.част.    | 1        | 1.109,04       | 1.109,04   |
| 831123153 | Рельс- 4 x 7.25 x 10 Оц   | 1        | 239,43         | 239,43     |
| 831123154 | Рельс - 4 x 7.25 x 20 О   | 11       | 360,39         | 3.964,29   |
|-----------|---------------------------|----------|----------------|------------|

Output (figures are given just for example):
|----------------|----------|----------------|--------|---------------|
| Price          | Discount | Sale           | Margin | Cost          |
| € 1 097 316,43 | 0%       | € 1 097 316,43 | 68%    | € 349 134,66  |
| € 1 097 316,43 | 2%       | € 1 075 370,10 | 68%    | € 349 134,66  |
| € 1 097 316,43 | 4%       | € 1 053 423,77 | 67%    | € 349 134,66  |
| € 1 097 316,43 | 6%       | € 1 031 477,44 | 66%    | € 349 134,66  |
| € 1 097 316,43 | 8%       | € 1 009 531,12 | 65%    | € 349 134,66  |
| € 1 097 316,43 | 10%      | € 987 584,79   | 65%    | € 349 134,66  |
| € 1 097 316,43 | 12%      | € 965 638,46   | 64%    | € 349 134,66  |
| € 1 097 316,43 | 14%      | € 943 692,13   | 63%    | € 349 134,66  |
|----------------|----------|----------------|--------|---------------|
