import numpy as np 
import pandas as pd 
from  sklearn.ensemble import RandomForestRegressor 
from sklearn.preprocessing import StandardScaler 
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score 
import xlwings as xw 

# Загрузка данных 
test = 'TEST.xlsx' 
train = 'TRAIN.xlsx' 
test_data = pd.read_excel(test, header=None, names=None) 
test_data.set_axis(test_data.iloc[0], axis=1, inplace=True) 
test_data = test_data[1:] 
train_data = pd.read_excel(train, header=None, names=None) 
train_data.set_axis(train_data.iloc[0], axis=1, inplace=True) 
train_data = train_data[1:] 
result = test_data['Item'].copy() 
result.reset_index() 
test_data = test_data.drop('Item', axis=1) 

# Оздание обучающей и тестовой выборок 
X_train = train_data.iloc[:, :-1].values 
y_train = train_data['OF_​PresME10_​Wq_​H'].values 

X_test = test_data.values 
# Масштабирование признаков 
scaler = StandardScaler() 
X_train_scaled = scaler.fit_transform(X_train) 
X_test_scaled = scaler.transform(X_test) 

# Обучение модели множественной регрессии 
regressor = RandomForestRegressor() 
regressor.fit(X_train_scaled, y_train) 

# Предсказание на тестовых данных 
y_pred = regressor.predict(X_test_scaled) 
y_test_pred = regressor.predict(X_train_scaled) 

# Создание нового датасета с индексом и столбцом 'predict' 
result = pd.DataFrame({'Item': result, 'predict': np.round(y_pred, 5)}) 
print(result) 
result = result.drop('Item', axis=1)

# Сохранение в формате CSV без столбца индексов 
result.to_csv('Molotochki_2.csv', index=False) 

# вывод данных в отдельный лист документа well_coord.xlsx 
print() 
print('Полная версия таблицы сохранена на листе "PREDICTION" документа TEST.xlsx') 
print() 
sheet_df_mapping = {"PREDICTION": result} 
with xw.App(visible=False) as app: 
    wb = app.books.open(test) 
    current_sheets = [sheet.name for sheet in wb.sheets] 
    for sheet_name in sheet_df_mapping.keys(): 
        if sheet_name in current_sheets: 
            wb.sheets(sheet_name).range("A1").value = sheet_df_mapping.get(sheet_name) 
        else: 
            new_sheet = wb.sheets.add(after=wb.sheets.count) 
            new_sheet.range("A1").value = sheet_df_mapping.get(sheet_name) 
            new_sheet.name = sheet_name 
    wb.save() 
    wb.close() 
# Сохранение данных на отдельном листе документа TEST.xlsx может быть полезным шагом потому, что исходя из данных OF, 
# можно говорить о том, можно ли использовать данные параметры при моделировании пласта, что создает перспективы для будущей автоматизации процесса 

# Расчет метрик 
mse = mean_squared_error(y_train, y_test_pred) 
mae = mean_absolute_error(y_train, y_test_pred) 
r2 = r2_score(y_train, y_test_pred) 

print("Среднеквадратическая ошибка (MSE):", mse) 
print("Средняя абсолютная ошибка (MAE):", mae) 
print("Коэффициент детерминации (R^2):", r2) 
print()
input("Нажмите Enter, чтобы закрыть программу.")
