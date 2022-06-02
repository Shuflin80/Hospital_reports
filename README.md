# Processing excel tables

The aim of my work is to create the automated validation mechanism for the hospital reports from the excel sheets in the required format


```python
import pandas as pd
```


```python
def find_nums(df1, df2, to_populate):
    """Identifier function to validate names in the reports"""
    
    df = pd.DataFrame(index=to_populate, columns=['Both', 'Not in 1st', "Not in 2nd"])
    for i in to_populate:
        to_insert = [0,0,0]
        
        rep1_names = list(df1[df1['Медицинская организация'].str.lower() == i.lower()]['ФИО'])
        rep2_names = list(df2[df2['Юридическое лицо2'].str.lower() == i.lower()]["ФИО"])
        
        check_1 = tuple(rep1_names)

        for name in check_1:
            if name in rep2_names:
                to_insert[0] += 1
                rep1_names.pop(rep1_names.index(name))
                rep2_names.pop(rep2_names.index(name))
            else:
                to_insert[2] += 1
                rep1_names.pop(rep1_names.index(name))

        to_insert[1] += len(rep2_names)
        
        
        df.loc[i] = to_insert
    return df
```


```python
path = 'Hospital_data.xlsx'

def solution(path):
    rep1 = pd.read_excel(path, sheet_name = "Отчет 1")
    rep2 = pd.read_excel(path, sheet_name = "Отчет 2")
    ref = pd.read_excel(path, sheet_name = "Справочник")

    mapping1 = dict(zip(ref['Отчет 1'], ref["С расшифровкой для свода"]))
    to_populate = pd.read_excel(path, sheet_name = "Получить свод")
    to_populate = to_populate = to_populate.iloc[6:, 0]
    to_populate.index.rename('entity')

    rep1 = rep1.iloc[:-1, [1,3]]
    rep1["Медицинская организация"] = rep1["Медицинская организация"].map(mapping1).fillna(rep1["Медицинская организация"])
    rep1 = rep1[rep1['Медицинская организация'].isin(to_populate)]
    rep1_patients = rep1.groupby('Медицинская организация')["ФИО"].count()

    df = pd.DataFrame(data=rep1_patients, index=to_populate)

    df.rename(columns={'ФИО': 'report_1'}, inplace=True)

    df.loc['ФГБУ ФКЦ ВМТ ФМБА РОССИИ', 'report_1'] = df.loc['ФГБУ ФКЦ ВМТ ФМБА России', 'report_1']

    rep1['ФИО'] = rep1['ФИО'].apply(lambda x: str(x).title())

    rep2['ФИО'] = rep2[["Фамилия", "Имя", "Отчество"]].astype(str).applymap(lambda x: x.title()).agg(" ".join, axis=1)
    rep2 = rep2.drop(["Фамилия", "Имя", "Отчество"], axis=1)

    rep2["Юридическое лицо2"] = rep2['Юридическое лицо'].replace({'ФГБУ "ФЕДЕРАЛЬНЫЙ КЛИНИЧЕСКИЙ ЦЕНТР ВЫСОКИХ МЕДИЦИНСКИХ ТЕХНОЛОГИЙ ФЕДЕРАЛЬНОГО МЕДИКО-БИОЛОГИЧЕСКОГО АГЕНТСТВА"': "ФГБУ ФКЦ ВМТ ФМБА России",
                                                                  "Королевская": "Королёвская", 'ГБУЗ МО Красногорская городская больница №': 'ГБУЗ МО Красногорская ГБ № ', 'ГБУЗ МОСКОВСКОЙ ОБЛАСТИ  "ПОДОЛЬСКАЯ ГОРОДСКАЯ КЛИНИЧЕСКАЯ БОЛЬНИЦА"': "ГБУЗ МО ПОДОЛЬСКАЯ ГКБ", "(?i)московская областная больница": "МОБ",'"': '', "имени": "им.",
                                                                  "областная клиническая больница": "ОКБ", "областная больница": "ОБ", "Московский областной научно-исследовательский клинический институт": "МОНИКИ",
                                                                  "Московская областная больница": "МОБ","первая районная больница": "ПРБ", "центральная районная клиническая больница":"ЦРКБ",
                                                                  "районная клиническая больница": "РКБ","МОСКОВСКОЙ ОБЛАСТИ": "МО", "детская городская больница": "ДГБ",
                                                                  "(?i)городская клиническая больница": "ГКБ","центральная районная больница": "ЦРБ", "(?i)районная больница": "РБ",
                                                                  "(?i)центральная городская больница": "ЦГБ", "(?i)ГОРОДСКАЯ БОЛЬНИЦА": "ГБ"}, regex=True)

    rep2 = rep2[rep2["Юридическое лицо2"].str.lower().isin(to_populate.str.lower())]
    rep2_patients = rep2.groupby(["Юридическое лицо2"])['ФИО'].count()

    df = df.merge(rep2_patients, left_on=df.index.str.lower(), right_on= rep2_patients.index.str.lower(), how='left').set_index(to_populate)
    df.drop('key_0', axis=1, inplace=True)
    df.rename(columns={'ФИО': 'report_2'}, inplace=True)

    df['DIFF'] = df.loc[:, 'report_2'].fillna(0) - df.loc[:, 'report_1'].fillna(0)

    diff_names = find_nums(rep1, rep2, to_populate)

    final = df.merge(diff_names, left_index=True, right_index=True)

    final['CHECK'] = final['DIFF'] - (final['Not in 1st'] - final['Not in 2nd'])
    a =pd.concat((pd.DataFrame(final.sum(axis=0).to_dict(), index=['Всего']), final))

    a.rename(columns = {'report_1': 'Пациенты на лечении по данным отчета 1', 'report_2': 'Пациенты на лечении по данным отчета 2',
                        'DIFF': 'Разница отчета 1 и отчета 2', 'Both': 'Есть в Отчете 1 и в Отчете 2 (совпадения)',
                        'Not in 1st': 'Есть в Отчете 2, нет в Отчете 1', 'Not in 2nd': 'Есть в отчете 1, нет в Отчете 2',
                        'CHECK': 'Проверка разницы'}, inplace=True)
    a.index.name = 'Наименование ЛПУ'
    a.to_excel("output.xlsx", na_rep='#N/A')
    
    return a

```


```python
pd.set_option('display.max_rows', 70)
```


```python
solution(path=path)
```




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
      <th>Пациенты на лечении по данным отчета 1</th>
      <th>Пациенты на лечении по данным отчета 2</th>
      <th>Разница отчета 1 и отчета 2</th>
      <th>Есть в Отчете 1 и в Отчете 2 (совпадения)</th>
      <th>Есть в Отчете 2, нет в Отчете 1</th>
      <th>Есть в отчете 1, нет в Отчете 2</th>
      <th>Проверка разницы</th>
    </tr>
    <tr>
      <th>Наименование ЛПУ</th>
      <th></th>
      <th></th>
      <th></th>
      <th></th>
      <th></th>
      <th></th>
      <th></th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>Всего</th>
      <td>5528.0</td>
      <td>6250.0</td>
      <td>722.0</td>
      <td>3931</td>
      <td>2319</td>
      <td>1597</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГАУЗ МО КЦВМиР</th>
      <td>409.0</td>
      <td>NaN</td>
      <td>-409.0</td>
      <td>0</td>
      <td>0</td>
      <td>409</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГАУЗ МО ХИМКИНСКАЯ ОБ</th>
      <td>573.0</td>
      <td>613.0</td>
      <td>40.0</td>
      <td>504</td>
      <td>109</td>
      <td>69</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО БАЛАШИХИНСКАЯ ОБ</th>
      <td>172.0</td>
      <td>214.0</td>
      <td>42.0</td>
      <td>156</td>
      <td>58</td>
      <td>16</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ВИДНОВСКАЯ РКБ</th>
      <td>431.0</td>
      <td>417.0</td>
      <td>-14.0</td>
      <td>358</td>
      <td>59</td>
      <td>73</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ДАВЫДОВСКАЯ РБ</th>
      <td>56.0</td>
      <td>52.0</td>
      <td>-4.0</td>
      <td>49</td>
      <td>3</td>
      <td>7</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ДМИТРОВСКАЯ ОБ</th>
      <td>67.0</td>
      <td>93.0</td>
      <td>26.0</td>
      <td>60</td>
      <td>33</td>
      <td>7</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ДОМОДЕДОВСКАЯ ЦГБ</th>
      <td>90.0</td>
      <td>107.0</td>
      <td>17.0</td>
      <td>73</td>
      <td>34</td>
      <td>17</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ЕГОРЬЕВСКАЯ ЦРБ</th>
      <td>94.0</td>
      <td>150.0</td>
      <td>56.0</td>
      <td>84</td>
      <td>66</td>
      <td>10</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ЖУКОВСКАЯ ГКБ</th>
      <td>95.0</td>
      <td>89.0</td>
      <td>-6.0</td>
      <td>78</td>
      <td>11</td>
      <td>17</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ИВАНТЕЕВСКАЯ ЦГБ</th>
      <td>73.0</td>
      <td>82.0</td>
      <td>9.0</td>
      <td>70</td>
      <td>12</td>
      <td>3</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО КАШИРСКАЯ ЦРБ</th>
      <td>49.0</td>
      <td>53.0</td>
      <td>4.0</td>
      <td>45</td>
      <td>8</td>
      <td>4</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО КЛИМОВСКАЯ ГБ №2</th>
      <td>27.0</td>
      <td>109.0</td>
      <td>82.0</td>
      <td>20</td>
      <td>89</td>
      <td>7</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО КОЛОМЕНСКАЯ ЦРБ</th>
      <td>108.0</td>
      <td>95.0</td>
      <td>-13.0</td>
      <td>91</td>
      <td>4</td>
      <td>17</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО КРАСНОГОРСКАЯ ГБ № 1</th>
      <td>135.0</td>
      <td>133.0</td>
      <td>-2.0</td>
      <td>124</td>
      <td>9</td>
      <td>11</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО КРАСНОГОРСКАЯ ГБ № 2</th>
      <td>5.0</td>
      <td>16.0</td>
      <td>11.0</td>
      <td>0</td>
      <td>16</td>
      <td>5</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ЛЕВОБЕРЕЖНАЯ ГП</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>0.0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ЛОБНЕНСКАЯ ЦГБ</th>
      <td>43.0</td>
      <td>39.0</td>
      <td>-4.0</td>
      <td>36</td>
      <td>3</td>
      <td>7</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ЛЬВОВСКАЯ РБ</th>
      <td>24.0</td>
      <td>33.0</td>
      <td>9.0</td>
      <td>13</td>
      <td>20</td>
      <td>11</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО СОЛНЕЧНОГОРСКАЯ ОБ</th>
      <td>341.0</td>
      <td>373.0</td>
      <td>32.0</td>
      <td>305</td>
      <td>68</td>
      <td>36</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО МОБ ИМ. ПРОФ. РОЗАНОВА В.Н.</th>
      <td>78.0</td>
      <td>112.0</td>
      <td>34.0</td>
      <td>68</td>
      <td>44</td>
      <td>10</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО МОЖАЙСКАЯ ЦРБ</th>
      <td>83.0</td>
      <td>70.0</td>
      <td>-13.0</td>
      <td>47</td>
      <td>23</td>
      <td>36</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО МОКПТД</th>
      <td>31.0</td>
      <td>NaN</td>
      <td>-31.0</td>
      <td>0</td>
      <td>0</td>
      <td>31</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО МОНИКИ ИМ. М.Ф. ВЛАДИМИРСКОГО</th>
      <td>NaN</td>
      <td>29.0</td>
      <td>29.0</td>
      <td>0</td>
      <td>29</td>
      <td>0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО МЫТИЩИНСКАЯ ГКБ</th>
      <td>513.0</td>
      <td>468.0</td>
      <td>-45.0</td>
      <td>366</td>
      <td>102</td>
      <td>147</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО НАРО-ФОМИНСКАЯ ОБ</th>
      <td>117.0</td>
      <td>123.0</td>
      <td>6.0</td>
      <td>103</td>
      <td>20</td>
      <td>14</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО НОГИНСКАЯ ЦРБ</th>
      <td>136.0</td>
      <td>130.0</td>
      <td>-6.0</td>
      <td>123</td>
      <td>7</td>
      <td>13</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ОДИНЦОВСКАЯ ОБ</th>
      <td>163.0</td>
      <td>253.0</td>
      <td>90.0</td>
      <td>97</td>
      <td>156</td>
      <td>66</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ОРЕХОВО-ЗУЕВСКАЯ ЦГБ</th>
      <td>45.0</td>
      <td>58.0</td>
      <td>13.0</td>
      <td>43</td>
      <td>15</td>
      <td>2</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ПАВЛОВО-ПОСАДСКАЯ ЦРБ</th>
      <td>42.0</td>
      <td>51.0</td>
      <td>9.0</td>
      <td>40</td>
      <td>11</td>
      <td>2</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ПОДОЛЬСКАЯ ГКБ</th>
      <td>61.0</td>
      <td>192.0</td>
      <td>131.0</td>
      <td>18</td>
      <td>174</td>
      <td>43</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ПОДОЛЬСКАЯ ДГБ</th>
      <td>22.0</td>
      <td>20.0</td>
      <td>-2.0</td>
      <td>12</td>
      <td>8</td>
      <td>10</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО СЕРГИЕВО-ПОСАДСКАЯ РБ</th>
      <td>361.0</td>
      <td>391.0</td>
      <td>30.0</td>
      <td>321</td>
      <td>70</td>
      <td>40</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО СЕРПУХОВСКАЯ ЦРБ</th>
      <td>95.0</td>
      <td>87.0</td>
      <td>-8.0</td>
      <td>76</td>
      <td>11</td>
      <td>19</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО СТУПИНСКАЯ ОКБ</th>
      <td>64.0</td>
      <td>80.0</td>
      <td>16.0</td>
      <td>60</td>
      <td>20</td>
      <td>4</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ШАТУРСКАЯ ЦРБ</th>
      <td>75.0</td>
      <td>85.0</td>
      <td>10.0</td>
      <td>72</td>
      <td>13</td>
      <td>3</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ЭЛЕКТРОСТАЛЬСКАЯ ЦГБ</th>
      <td>123.0</td>
      <td>129.0</td>
      <td>6.0</td>
      <td>93</td>
      <td>36</td>
      <td>30</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ФГБОУ ВО МГМСУ ИМ. А.И. ЕВДОКИМОВА МИНЗДРАВА РОССИИ</th>
      <td>NaN</td>
      <td>647.0</td>
      <td>647.0</td>
      <td>0</td>
      <td>647</td>
      <td>0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ФГБУ ФКЦ ВМТ ФМБА РОССИИ</th>
      <td>108.0</td>
      <td>67.0</td>
      <td>-41.0</td>
      <td>0</td>
      <td>67</td>
      <td>108</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>БПНЦ РАН</th>
      <td>1.0</td>
      <td>NaN</td>
      <td>-1.0</td>
      <td>0</td>
      <td>0</td>
      <td>1</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>Больница НЦЧ РАН</th>
      <td>4.0</td>
      <td>NaN</td>
      <td>-4.0</td>
      <td>0</td>
      <td>0</td>
      <td>4</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГАУЗ МО КЛИНСКАЯ ГБ</th>
      <td>NaN</td>
      <td>1.0</td>
      <td>1.0</td>
      <td>0</td>
      <td>1</td>
      <td>0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГАУЗ МО ЦГБ им. М.В. Гольца</th>
      <td>2.0</td>
      <td>8.0</td>
      <td>6.0</td>
      <td>1</td>
      <td>7</td>
      <td>1</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ВОЛОКОЛАМСКАЯ ЦРБ</th>
      <td>16.0</td>
      <td>2.0</td>
      <td>-14.0</td>
      <td>0</td>
      <td>2</td>
      <td>16</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ВОСКРЕСЕНСКАЯ ПРБ</th>
      <td>4.0</td>
      <td>10.0</td>
      <td>6.0</td>
      <td>3</td>
      <td>7</td>
      <td>1</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ГОЛИЦЫНСКАЯ ПОЛИКЛИНИКА</th>
      <td>1.0</td>
      <td>NaN</td>
      <td>-1.0</td>
      <td>0</td>
      <td>0</td>
      <td>1</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ДЗЕРЖИНСКАЯ ГБ</th>
      <td>1.0</td>
      <td>10.0</td>
      <td>9.0</td>
      <td>1</td>
      <td>9</td>
      <td>0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ДОЛГОПРУДНЕНСКАЯ ЦГБ</th>
      <td>NaN</td>
      <td>23.0</td>
      <td>23.0</td>
      <td>0</td>
      <td>23</td>
      <td>0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО КОРОЛЁВСКАЯ ГБ</th>
      <td>12.0</td>
      <td>27.0</td>
      <td>15.0</td>
      <td>10</td>
      <td>17</td>
      <td>2</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ЛОТОШИНСКАЯ ЦРБ</th>
      <td>1.0</td>
      <td>1.0</td>
      <td>0.0</td>
      <td>1</td>
      <td>0</td>
      <td>0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ЛЫТКАРИНСКАЯ ГБ</th>
      <td>1.0</td>
      <td>2.0</td>
      <td>1.0</td>
      <td>0</td>
      <td>2</td>
      <td>1</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ЛЮБЕРЕЦКАЯ ОБ</th>
      <td>398.0</td>
      <td>414.0</td>
      <td>16.0</td>
      <td>301</td>
      <td>113</td>
      <td>97</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО РАМЕНСКАЯ ЦРБ</th>
      <td>2.0</td>
      <td>1.0</td>
      <td>-1.0</td>
      <td>0</td>
      <td>1</td>
      <td>2</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО СЕРЕБРЯНО-ПРУДСКАЯ ЦРБ</th>
      <td>2.0</td>
      <td>4.0</td>
      <td>2.0</td>
      <td>2</td>
      <td>2</td>
      <td>0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО СТУПИНСКАЯ ЦРКБ</th>
      <td>2.0</td>
      <td>NaN</td>
      <td>-2.0</td>
      <td>0</td>
      <td>0</td>
      <td>2</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ТАЛДОМСКАЯ ЦРБ</th>
      <td>11.0</td>
      <td>14.0</td>
      <td>3.0</td>
      <td>6</td>
      <td>8</td>
      <td>5</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ГБУЗ МО ЧЕХОВСКАЯ ОБ</th>
      <td>1.0</td>
      <td>5.0</td>
      <td>4.0</td>
      <td>1</td>
      <td>4</td>
      <td>0</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ФБУЗ МСЧ №9 ФМБА России</th>
      <td>2.0</td>
      <td>NaN</td>
      <td>-2.0</td>
      <td>0</td>
      <td>0</td>
      <td>2</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ФГБУ "3 центральный военный клинический госпиталь имени А.А. Вишневского" Минобороны России</th>
      <td>8.0</td>
      <td>NaN</td>
      <td>-8.0</td>
      <td>0</td>
      <td>0</td>
      <td>8</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ФГБУ ФКЦ ВМТ ФМБА России</th>
      <td>108.0</td>
      <td>67.0</td>
      <td>-41.0</td>
      <td>0</td>
      <td>67</td>
      <td>108</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ФГБУЗ МСЧ № 154 ФМБА России</th>
      <td>2.0</td>
      <td>NaN</td>
      <td>-2.0</td>
      <td>0</td>
      <td>0</td>
      <td>2</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ФГБУЗ МСЧ № 164 ФМБА России</th>
      <td>1.0</td>
      <td>NaN</td>
      <td>-1.0</td>
      <td>0</td>
      <td>0</td>
      <td>1</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ФГБУЗ МСЧ №152 ФМБА России</th>
      <td>27.0</td>
      <td>1.0</td>
      <td>-26.0</td>
      <td>0</td>
      <td>1</td>
      <td>27</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>ФГБУЗ ЦМСЧ № 94 ФМБА России</th>
      <td>11.0</td>
      <td>NaN</td>
      <td>-11.0</td>
      <td>0</td>
      <td>0</td>
      <td>11</td>
      <td>0.0</td>
    </tr>
    <tr>
      <th>Филиал № 1ФГБУ «3 ЦВКГ им. А. А. Вишневского» Минобороны России</th>
      <td>1.0</td>
      <td>NaN</td>
      <td>-1.0</td>
      <td>0</td>
      <td>0</td>
      <td>1</td>
      <td>0.0</td>
    </tr>
  </tbody>
</table>
</div>



    1. Для Отчета 1: Отсутствующие значения для медицинских учреждений обусловлены в большинстве случаев отсутствием данных по этим местам в самом отчете. Однако, некоторые больницы например (Клинская ГБ) имеет неправильное определение в справочнике ("Клинская ОБ"). Также в своде присутствует дубликат: "ФГБУ ФКЦ ВМТ ФМБА РОССИИ" и "ФГБУ ФКЦ ВМТ ФМБА России". Так как это содержиться в форме отчета, я оставил оба этих названия и данные этих двух заведений полностью копируют друг друга.

    2. Для Отчета 2: Определения в справочнике не подходят ни под одно значение лечебных учреждений в отчете! Я вручную подогнал названия ЛПУ с требуемыми именами в своде. Все отсутствующие значения говорят об отсутствии данных в отчете.

    3. Я специально использовал функцию, которая будет считать пациентов с одинаковым ФИО как двух разных пациентов, так как не указано обратного. Таким образом в проверке нет несоответствий. В инном случае, погрешность составляла от 1 до 3 пациентов (С одинаковым ФИО).

    4. Функция экспортирует получившуюся таблицу в xlsx формате


```python

```
