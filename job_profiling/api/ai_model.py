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
from api import utils


class DataModel:
    def __init__(self, data_path, columns):  # columns is a list of columns names as appeared in the training file
        self.data = pd.read_csv(data_path)[columns]

    def build_and_train_model(self):
        self.build_encoders()

        X = self.data.drop(utils.TARGET, axis=1)
        y = self.data[utils.TARGET]
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.15)

        model = DecisionTreeClassifier(max_depth=8)
        model.fit(X_train, y_train)

        train_predictions = model.predict(X_train)

        test_predictions = model.predict(X_test)

        train_evaluation = classification_report(y_train, train_predictions)
        print(train_evaluation)

        test_evaluation = classification_report(y_test, test_predictions)
        print(test_evaluation)

        joblib.dump(model, os.path.join(settings.MODELS_PATH, 'model.joblib'))
        return model

    def build_encoders(self):
        for feature in utils.CATEGORICAL_FEATURES:
            if feature in utils.ONE_HOT_ENCODED_FEATURES:
                encoder = OneHotEncoder()
                temp = pd.DataFrame(
                    encoder.fit_transform(self.data[[feature]]).toarray(),
                    columns=[name.replace(f'{feature}_', '' ) for name in encoder.get_feature_names_out()]
                )
                print(temp.head())
                self.data = pd.concat([self.data, temp], axis=1).drop(feature, axis=1)

            elif feature in utils.LABEL_ENCODED_FEATURES:
                encoder = LabelEncoder()
                self.data[feature] = encoder.fit_transform(self.data[feature])
            else:
                raise ValueError(f"'feature' can't be encoded.")

            # print(self.data[feature])
            # print(feature, '-' * 100, )
            # consts.ENCODERS[feature] = encoder

            joblib.dump(encoder, os.path.join(settings.MODELS_PATH, f'{feature}_encoder.joblib'))

        print(self.data.head())