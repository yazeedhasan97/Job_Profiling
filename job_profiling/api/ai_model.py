import os.path
import sys

import pandas as pd
from sklearn.tree import DecisionTreeClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, classification_report
from sklearn.preprocessing import LabelEncoder, OneHotEncoder
import joblib
import matplotlib.pyplot as plt
import re
from django.conf import settings
import numpy as np
from api import consts


class DataModel:
    def __init__(self, data_path, columns):  # columns is a list of columns names as appeared in the training file
        self.data = pd.read_csv(data_path)[columns]
        # self.data = self.unify_columns_names(pd.read_csv(data_path)[columns])
        # print(self.data.head(10))
        # self.clean_na_values()
        # print(self.data.info())
        # self.data = self.unify_categorical_values(self.data)
        # print(self.data.tail(10))

    def unify_columns_names(self, data):
        cols = []
        for col in data.columns:
            col = re.sub(r"(\w)([A-Z])", r"\1 \2", col)
            cols.append(col.strip(' .()[]{}/\#@*^!?').replace('_', ' ').replace(' ', '_').lower())
        data.columns = cols
        print(data.head())
        return data

    def build_and_train_model(self):
        self.build_encoders()

        X = self.data.drop(consts.TARGET, axis=1)
        y = self.data[consts.TARGET]
        X_train, X_test, y_train, y_test = train_test_split(X, y)

        model = DecisionTreeClassifier(max_depth=6)
        model.fit(X_train, y_train)

        train_predictions = model.predict(X_train)

        test_predictions = model.predict(X_test)

        train_evaluation = classification_report(y_train, train_predictions)
        print(train_evaluation)

        test_evaluation = classification_report(y_test, test_predictions)
        print(test_evaluation)

        joblib.dump(model, os.path.join(settings.BASE_DIR, 'PreProcessingAndModelCode', 'runs', 'model.joblib'))
        return model

    # def clean_na_values(self):
    #     for col in self.data.columns:
    #         if col in consts.FILL_NA_TEMPLATE:
    #             self.data[col] = self.data[col].fillna(consts.FILL_NA_TEMPLATE[col])
    #         else:
    #             self.data[col] = self.data[col].fillna(consts.FILL_NA_TEMPLATE[consts.OTHERS])

    # def unify_categorical_values(self, data):
    #     for col in consts.CATEGORICAL_FEATURES:
    #         data[col] = data[col].str.replace(' ', '_')
    #         data[col] = data[col].str.replace('-', '_').str.lower()
    #     return data

    # def overview(self, data):
    #     total_columns_onehot = 0
    #     total_columns_label = 0
    #     for col in data.columns:
    #         unq = data[col].unique()
    #         total_columns_onehot += len(unq)
    #         total_columns_label += 1
    #         print(f'The column {col} has {len(unq)} unique values separated as: {unq}')
    #         data[col].value_counts().plot(kind='bar')
    #         plt.show()
    #     print(f'total_columns_onehot = {total_columns_onehot}')
    #     print(f'total_columns_label = {total_columns_label}')

    def build_encoders(self):
        for feature in consts.CATEGORICAL_FEATURES:
            if feature in consts.ONE_HOT_ENCODED_FEATURES:
                encoder = OneHotEncoder()
                temp = pd.DataFrame(
                    encoder.fit_transform(self.data[[feature]]).toarray(),
                    columns=[name.replace(f'{feature}_', '' ) for name in encoder.get_feature_names_out()]
                )
                print(temp.head())
                self.data = pd.concat([self.data, temp], axis=1).drop(feature, axis=1)

            elif feature in consts.LABEL_ENCODED_FEATURES:
                encoder = LabelEncoder()
                self.data[feature] = encoder.fit_transform(self.data[feature])
            else:
                raise ValueError(f"'feature' can't be encoded.")

            # print(self.data[feature])
            # print(feature, '-' * 100, )
            # consts.ENCODERS[feature] = encoder
            print(self.data.head())
            joblib.dump(encoder, os.path.join(settings.BASE_DIR, 'PreProcessingAndModelCode', 'runs', f'{feature}_encoder.joblib'))