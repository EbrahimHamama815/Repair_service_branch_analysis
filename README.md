# Repair_service_branch_analysis
A Yearly Analysis of the revenue of a multi-branch mobile rapair shop.


In this project we need to analyze and compare revenue of 2024 and 2025 for a large mobile repair shop with +40 branches.

The real challenge here is the data structure, where we have monthly revenue on an excel workbook and monthly accepted/rejected orders on a totaly different workbook, not only that but each workbook is divided by branches, each branch having a seperate sheet, each sheet formatted like this:

<img width="1091" height="570" alt="image" src="https://github.com/user-attachments/assets/66cdb496-7d1b-46db-9df3-980a98913ef5" />

<hr>

<img width="1128" height="569" alt="image" src="https://github.com/user-attachments/assets/39c8bcf1-a8b3-49e2-b1cc-1cffeafef905" />

<p>
<br>
</p>

So, in order to perform any kind of analysis we first need to prepare the data. And for this type of messy data, Python is the perfect tool for it.


## Preparing the Data

The goal here is to combine all branches sheets in one table so we can process it in Power BI.

Lets go through the first workbook, we import the whole workbook as dictionary of dataframes and specifiy the needed sheets only (because it has a total sheet).

```python
# get needed sheets only
stores = list(pd.read_excel("تصنيف العملاء_2025-2024(1).xlsx",sheet_name=None).keys())[:-2]

# load all dataframes in a dictionary
df_dict = pd.read_excel("تصنيف العملاء_2025-2024(1).xlsx",sheet_name=stores)
```

And then we perfrom a loop on every sheet and apply the needed cleaning as explained in the code comments:

<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Unnamed: 0</th>
      <th>2024</th>
      <th>Unnamed: 2</th>
      <th>Unnamed: 3</th>
      <th>2025</th>
      <th>Unnamed: 5</th>
      <th>Unnamed: 6</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>month</td>
      <td>عدد الفواتير</td>
      <td>محقق شهريا</td>
      <td>متوسط الفاتورة</td>
      <td>عدد الفواتير</td>
      <td>محقق شهريا</td>
      <td>متوسط الفاتورة</td>
    </tr>
    <tr>
      <th>1</th>
      <td>Jan</td>
      <td>101</td>
      <td>127651</td>
      <td>1263.871287</td>
      <td>95</td>
      <td>120568</td>
      <td>1269.136842</td>
    </tr>
    <tr>
      <th>2</th>
      <td>Feb</td>
      <td>77</td>
      <td>139703</td>
      <td>1814.324675</td>
      <td>93</td>
      <td>160850</td>
      <td>1729.569892</td>
    </tr>
    <tr>
      <th>3</th>
      <td>Mar</td>
      <td>82</td>
      <td>130673</td>
      <td>1593.573171</td>
      <td>78</td>
      <td>93050</td>
      <td>1192.948718</td>
    </tr>
    <tr>
      <th>4</th>
      <td>Apr</td>
      <td>80</td>
      <td>116850</td>
      <td>1460.625</td>
      <td>81</td>
      <td>162250</td>
      <td>2003.08642</td>
    </tr>
    <tr>
      <th>5</th>
      <td>May</td>
      <td>88</td>
      <td>117828</td>
      <td>1338.954545</td>
      <td>64</td>
      <td>97832</td>
      <td>1528.625</td>
    </tr>
    <tr>
      <th>6</th>
      <td>Jun</td>
      <td>101</td>
      <td>135266</td>
      <td>1339.267327</td>
      <td>61</td>
      <td>66330</td>
      <td>1087.377049</td>
    </tr>
    <tr>
      <th>7</th>
      <td>Jul</td>
      <td>74</td>
      <td>118312.5</td>
      <td>1598.817568</td>
      <td>110</td>
      <td>149150</td>
      <td>1355.909091</td>
    </tr>
    <tr>
      <th>8</th>
      <td>Aug</td>
      <td>64</td>
      <td>110191.5</td>
      <td>1721.742188</td>
      <td>111</td>
      <td>151850</td>
      <td>1368.018018</td>
    </tr>
    <tr>
      <th>9</th>
      <td>Sep</td>
      <td>91</td>
      <td>94200</td>
      <td>1035.164835</td>
      <td>88</td>
      <td>105300</td>
      <td>1196.590909</td>
    </tr>
    <tr>
      <th>10</th>
      <td>Oct</td>
      <td>109</td>
      <td>156809</td>
      <td>1438.614679</td>
      <td>62</td>
      <td>69750</td>
      <td>1125</td>
    </tr>
    <tr>
      <th>11</th>
      <td>nov</td>
      <td>86</td>
      <td>127200</td>
      <td>1479.069767</td>
      <td>57</td>
      <td>104400</td>
      <td>1831.578947</td>
    </tr>
    <tr>
      <th>12</th>
      <td>Dec</td>
      <td>58</td>
      <td>68550</td>
      <td>1181.896552</td>
      <td>71</td>
      <td>99900</td>
      <td>1407.042254</td>
    </tr>
    <tr>
      <th>13</th>
      <td>Total</td>
      <td>1011</td>
      <td>1443234</td>
      <td>1427.531157</td>
      <td>971</td>
      <td>1381230</td>
      <td>1422.481977</td>
    </tr>
    <tr>
      <th>14</th>
      <td>A</td>
      <td>1801-5000</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>تنصيف العميل  B</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>15</th>
      <td>b</td>
      <td>1201-1800</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>16</th>
      <td>c</td>
      <td>600-1200</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>17</th>
      <td>NaN</td>
      <td>2024</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>2025</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>18</th>
      <td>Q1</td>
      <td>260</td>
      <td>398027</td>
      <td>1530.873077</td>
      <td>266</td>
      <td>374468</td>
      <td>1407.774436</td>
    </tr>
    <tr>
      <th>19</th>
      <td>Q2</td>
      <td>269</td>
      <td>369944</td>
      <td>1375.256506</td>
      <td>206</td>
      <td>326412</td>
      <td>1584.524272</td>
    </tr>
    <tr>
      <th>20</th>
      <td>Q3</td>
      <td>229</td>
      <td>322704</td>
      <td>1409.187773</td>
      <td>309</td>
      <td>406300</td>
      <td>1314.886731</td>
    </tr>
    <tr>
      <th>21</th>
      <td>Q4</td>
      <td>253</td>
      <td>352559</td>
      <td>1393.513834</td>
      <td>190</td>
      <td>274050</td>
      <td>1442.368421</td>
    </tr>
  </tbody>
</table>
</div>


```python


```

