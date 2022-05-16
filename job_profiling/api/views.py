import os.path

from rest_framework import status
from rest_framework.decorators import api_view
from django.http import JsonResponse

from api import ai_model
from api import utils
import joblib
import numpy as np
from django.conf import settings

# this reads the files on the first run of the API, and as long as
# the API is running and the files exists they wil work fine
# if the API is restarted and the models/encoder were not there it will autogenerate them
# but note, this will not work if the data file have been removed
if os.path.exists(utils.MODEL):
    model = joblib.load(utils.MODEL)
    utils.ENCODERS[utils.GENDER] = joblib.load(utils.ENCODERS[utils.GENDER])
    utils.ENCODERS[utils.EDUCATION] = joblib.load(utils.ENCODERS[utils.EDUCATION])
    utils.ENCODERS[utils.GOVERNORATE] = joblib.load(utils.ENCODERS[utils.GOVERNORATE])
    utils.ENCODERS[utils.DISABILITY] = joblib.load(utils.ENCODERS[utils.DISABILITY])
else:
    model = ai_model.DataModel(
        os.path.join(
            settings.DATA_PATH,
            'final_outputs.csv'
        ),
        utils.API_CONTRIBUTORS + utils.API_TARGET
    ).build_and_train_model()
    utils.ENCODERS[utils.GENDER] = joblib.load(utils.ENCODERS[utils.GENDER])
    utils.ENCODERS[utils.EDUCATION] = joblib.load(utils.ENCODERS[utils.EDUCATION])
    utils.ENCODERS[utils.GOVERNORATE] = joblib.load(utils.ENCODERS[utils.GOVERNORATE])
    utils.ENCODERS[utils.DISABILITY] = joblib.load(utils.ENCODERS[utils.DISABILITY])


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
    for feature in utils.API_CONTRIBUTORS:  # request to submit each parameter using its name
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
    for feature in utils.API_CONTRIBUTORS:  # accept parameter and preprocess them
        # print('Feature -------------------------- ', feature)
        if feature in utils.LABEL_ENCODED_FEATURES:
            # print('Feature Value-------------------------- ', data.get(feature))
            # na_fill_value = TEMPLATE_DATA_CATEGORIES[feature][NA_FILL_VALUE]
            # print('NA FILL -------------------------- ', na_fill_value)
            value = eval(utils.TEMPLATE_DATA_CATEGORIES[feature][utils.CODE].format('data.get(feature, None)'))
            # print('Preprocessed VALUE  -------------------------- ', value)
            model_input[feature] = int(utils.ENCODERS[feature].transform([value])[0]) if value is not None else value

        elif feature in utils.ONE_HOT_ENCODED_FEATURES:
            # categories = ENCODERS[feature].get_feature_names_out()
            # print(categories)
            # for i in range(len(categories)):
            #     model_input[categories[i]] = 0 if data.get(feature) not in categories[i] else 1
            # print('Feature Value-------------------------- ', data.get(feature))
            # na_fill_value = TEMPLATE_DATA_CATEGORIES[feature][NA_FILL_VALUE]
            # print('NA FILL -------------------------- ', na_fill_value)
            value = eval(utils.TEMPLATE_DATA_CATEGORIES[feature][utils.CODE].format('data.get(feature, None)'))
            # print('Preprocessed VALUE  -------------------------- ', value)
            if value is not None:
                item = utils.ENCODERS[feature].transform([[value]]).indices
                # print('Encoded Value -------------------------- ', item)
                # print(list(item))
                lst = {
                    v.replace(f'{feature}_', ''): 1 if item == i else 0 for i, v in
                    enumerate(utils.ENCODERS[feature].get_feature_names_out())
                }
            else:
                lst = [None]
            # print('Transferred Items Value -------------------------- ', lst)
            model_input.update(lst)
        else:
            # print(f'Feature {feature} Value-------------------------- ', data.get(feature))
            # na_fill_value = TEMPLATE_DATA_CATEGORIES[feature][NA_FILL_VALUE]
            # print('NA FILL -------------------------- ', na_fill_value)
            item = eval(utils.TEMPLATE_DATA_CATEGORIES[feature][utils.CODE].format('data.get(feature, None)'))
            # print('Preprocessed VALUE  -------------------------- ', item)
            model_input[feature] = int(item)

            # na_fill_value = TEMPLATE_DATA_CATEGORIES[feature][NA_FILL_VALUE]
            # item = int(data.get(feature, na_fill_value))
            # model_input[feature] = item

        # print('-' * 50, 'New Variable', '-' * 50)

    print('Model_Inputs', '-' * 50, )
    dd = list(model_input.values())
    if any(elem is None for elem in dd) or any(elem is np.nan for elem in dd):
        print('Went for missing -------------------------------')
        res = utils.predict_node_for_missing_value(model, dd)[-1]
    else:
        print('Went for complete -------------------------------')
        res = model.predict([dd])[0]
    print('RES :', res)
    return JsonResponse({'cluster': str(res)}, safe=False, status=status.HTTP_200_OK)
