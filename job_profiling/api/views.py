import os.path

from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response
from django.http import HttpRequest, JsonResponse

from api import ai_model
from api.consts import *
from api.documents import EvaluationOutput
from api.serializers import EvaluationOutputSerializer
import random
import joblib
import pandas as pd
from django.conf import settings

if os.path.exists(MODEL):
    model = joblib.load(MODEL)
    ENCODERS[GENDER] = joblib.load(ENCODERS[GENDER])
    ENCODERS[EDUCATION] = joblib.load(ENCODERS[EDUCATION])
    ENCODERS[GOVERNORATE] = joblib.load(ENCODERS[GOVERNORATE])
else:
    model = ai_model.DataModel(
        os.path.join(
            settings.BASE_DIR,
            'PreProcessingAndModelCode',
            'data',
            'fixed_tree_results',
            'prefinals.csv'
        ),
        CONTRIBUTORS
    ).build_and_train_model()
    ENCODERS[GENDER] = joblib.load(ENCODERS[GENDER])
    ENCODERS[EDUCATION] = joblib.load(ENCODERS[EDUCATION])
    ENCODERS[GOVERNORATE] = joblib.load(ENCODERS[GOVERNORATE])
print(model)
print(ENCODERS)


# URL example
# http://127.0.0.1:8000/cluster/?experience=0&age=20&education_level=2&employment_status=2&disability_no_disability=1&governorate_amman=1

# @api_view(['GET'])
# def info_list(request):
#     """
#
#     """
#
#     results = EvaluationOutput.objects.all()[:10]
#     serializer = EvaluationOutputSerializer(results, many=True)
#     return Response(serializer.data)


@api_view(['GET', 'POST'])
def get_cluster(request):
    parameters = []
    for feature in API_FEATURES:  # request to submit each parameter using its name
        parameters.append(request.query_params.get(feature))

    if request.method == 'GET':
        print(request.GET)
        data = request.GET
        print(data)
    else:
        print(request.POST)
        data = request.POST
        print(data)

    model_input = {}
    for feature in API_FEATURES:  # accept parameter and preprocess them
        print('Feature -------------------------- ', feature)
        if feature in LABEL_ENCODED_FEATURES:
            print('Feature Value-------------------------- ', data.get(feature))
            na_fill_value = TEMPLATE_DATA_CATEGORIES[feature][NA_FILL_VALUE]
            print('NA FILL -------------------------- ', na_fill_value)
            value = eval(TEMPLATE_DATA_CATEGORIES[feature][CODE].format('data.get(feature, na_fill_value)'))
            print('Preprocessed VALUE  -------------------------- ', value)
            item = ENCODERS[feature].transform([value])[0]
            print('Encoded Value -------------------------- ', item)
            model_input[feature] = int(item)

        elif feature in ONE_HOT_ENCODED_FEATURES:
            # categories = ENCODERS[feature].get_feature_names_out()
            # print(categories)
            # for i in range(len(categories)):
            #     model_input[categories[i]] = 0 if data.get(feature) not in categories[i] else 1
            print('Feature Value-------------------------- ', data.get(feature))
            na_fill_value = TEMPLATE_DATA_CATEGORIES[feature][NA_FILL_VALUE]
            print('NA FILL -------------------------- ', na_fill_value)
            value = eval(TEMPLATE_DATA_CATEGORIES[feature][CODE].format('data.get(feature, na_fill_value)'))
            print('Preprocessed VALUE  -------------------------- ', value)
            item = ENCODERS[feature].transform([[value]]).indices
            print('Encoded Value -------------------------- ', item)
            print(list(item))
            lst = {v.replace(f'{feature}_', '') : 1 if item == i else 0 for i, v in enumerate(ENCODERS[feature].get_feature_names_out())}
            print('Transferred Items Value -------------------------- ', lst)
            model_input.update(lst)
        else:
            print('Feature Value-------------------------- ', data.get(feature))
            na_fill_value = TEMPLATE_DATA_CATEGORIES[feature][NA_FILL_VALUE]
            print('NA FILL -------------------------- ', na_fill_value)
            item = eval(TEMPLATE_DATA_CATEGORIES[feature][CODE].format('data.get(feature, na_fill_value)'))
            print('Preprocessed VALUE  -------------------------- ', item)
            model_input[feature] = int(item)

            # na_fill_value = TEMPLATE_DATA_CATEGORIES[feature][NA_FILL_VALUE]
            # item = int(data.get(feature, na_fill_value))
            # model_input[feature] = item

        print('-' * 50, 'New Variable', '-' * 50)

    print('Model_Inputs', '-' * 50, model_input)
    res = model.predict(pd.DataFrame(model_input, index=[0]))
    print('RES :', res[0])
    return JsonResponse({'cluster': str(res[0])}, safe=False, status=status.HTTP_200_OK)
