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
# else:
#     model = ai_model.DataModel(
#         os.path.join(
#             settings.DATA_PATH,
#             'final_outputs.csv'
#         ),
#         utils.API_CONTRIBUTORS + utils.API_TARGET
#     ).build_and_train_model()
#     utils.ENCODERS[utils.GENDER] = joblib.load(utils.ENCODERS[utils.GENDER])
#     utils.ENCODERS[utils.EDUCATION] = joblib.load(utils.ENCODERS[utils.EDUCATION])
#     utils.ENCODERS[utils.GOVERNORATE] = joblib.load(utils.ENCODERS[utils.GOVERNORATE])
#     utils.ENCODERS[utils.DISABILITY] = joblib.load(utils.ENCODERS[utils.DISABILITY])


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
        # print(request.GET)
        data = request.GET
        # print(data)
    else:
        # print(request.POST)
        data = request.POST
        # print(data)

    model_input = {}
    for feature in utils.API_CONTRIBUTORS:  # accept parameter and preprocess them

        if feature in utils.LABEL_ENCODED_FEATURES:
            value = eval(utils.TEMPLATE_DATA_CATEGORIES[feature][utils.CODE].format('data.get(feature, None)'))
            model_input[feature] = int(utils.ENCODERS[feature].transform([value])[0]) if value is not None else value

        elif feature in utils.ONE_HOT_ENCODED_FEATURES:
            value = eval(utils.TEMPLATE_DATA_CATEGORIES[feature][utils.CODE].format('data.get(feature, None)'))
            if value is not None:
                item = utils.ENCODERS[feature].transform([[value]]).indices
                lst = {
                    v: 1 if i == item else 0 for i, v in
                    enumerate(utils.ENCODERS[feature].get_feature_names_out())
                }
            else:
                lst = {
                    v: 1 if 'amman' in v else 0 for i, v in
                    enumerate(utils.ENCODERS[feature].get_feature_names_out())
                }
                print()
            print(lst)
            model_input.update(lst)
            print(model_input)
        else:
            item = eval(utils.TEMPLATE_DATA_CATEGORIES[feature][utils.CODE].format('data.get(feature, None)'))
            model_input[feature] = int(item)

    print('Model_Inputs', '-' * 50, )
    dd = list(model_input.values())
    if any(elem is None or elem is np.nan for elem in dd):
        print('Went for missing -------------------------------')
        res = utils.predict_node_for_missing_value(model, dd)[-1]
    else:
        print('Went for complete -------------------------------')
        res = model.predict([dd])[0]
    print('RES :', res)
    return JsonResponse({'cluster': str(res)}, safe=False, status=status.HTTP_200_OK)
