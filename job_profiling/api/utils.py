from django.conf import settings
import os
import numpy as np

# Attributes and contributors
GOVERNORATE = 'governorate'
AGE = 'age'
EXPERIENCE = 'experience'
EDUCATION = 'education'
GENDER = 'gender'
DISABILITY = 'disability'

API_CONTRIBUTORS = [GOVERNORATE, AGE, EXPERIENCE, EDUCATION, GENDER, DISABILITY]

TARGET = 'clusters'
API_TARGET = [TARGET]


ONE_HOT_ENCODED_FEATURES = []
LABEL_ENCODED_FEATURES = [GENDER, EDUCATION, GOVERNORATE, DISABILITY]
CATEGORICAL_FEATURES = ONE_HOT_ENCODED_FEATURES + LABEL_ENCODED_FEATURES

ENCODERS = {
    GENDER: os.path.join(settings.MODELS_PATH, f'{GENDER}_encoder.joblib'),
    EDUCATION: os.path.join(settings.MODELS_PATH, f'{EDUCATION}_encoder.joblib'),
    GOVERNORATE: os.path.join(settings.MODELS_PATH, f'{GOVERNORATE}_encoder.joblib'),
    DISABILITY: os.path.join(settings.MODELS_PATH, f'{DISABILITY}_encoder.joblib'),
}
MODEL = os.path.join(settings.MODELS_PATH, 'model.joblib')

# Integrity strings
NA_FILL_VALUE = 'NA_FILL_VALUE'
CATEGORIES = 'CATEGORIES'
CODE = 'CODE'

# FIRST_JOB = 'first_ind'
# INDUSTRY = 'industry'
# EMPLOYMENT = 'employment'
# DISABILITY = 'disability'

TEMPLATE_DATA_CATEGORIES = {
    # FIRST_JOB: {
    #     CATEGORIES: [0, 1],
    #     CODE: "first_job_code({0})",
    #     NA_FILL_VALUE: 0
    # },

    # EMPLOYMENT: {
    #     CATEGORIES: ['unemployed', 'formal_worker', 'daily_worker', 'informal_worker', 'housewife', 'self_employed'],
    #     CODE: "employment_code({0})",
    #     NA_FILL_VALUE: 'unemployed'
    # },
    # INDUSTRY: {
    #     CATEGORIES: [
    #         'industry', 'construction', "tourism", "education",
    #         'wholesale_and_retail_trade', 'agriculture_hunting_and_fishery',
    #         'mining_and_quarrying', 'transportation_communication_and_storage', 'electricity_gas_and_water',
    #         'community_service_activities', 'financial_activities', 'health_and_social_activities',
    #         'real_estate_activities', 'public_administration_defense_and_social_security',
    #         'international_organizations', 'private_families_who_hire_individuals_for_household_chores',
    #     ],
    #     CODE: "industry_code({0})",
    #     NA_FILL_VALUE: 'none'
    # },
    DISABILITY: {
        CATEGORIES: ['with_disability', 'no_disability'],
        CODE: "utils.disability_code({0})",
        NA_FILL_VALUE: 'no_disability'
    },
    EXPERIENCE: {
        CATEGORIES: [0, 1, 5, 10, 15, 20],
        CODE: "utils.experience_code({0})",
        NA_FILL_VALUE: 0
    },
    AGE: {
        CATEGORIES: [10, 20, 30, 40, 50, 60],
        CODE: "utils.age_code({0})",
        NA_FILL_VALUE: 30
    },

    EDUCATION: {
        CATEGORIES: ['bachelor_or_above', 'vocational_training', 'middle_diploma', 'secondary_or_below'],
        CODE: "utils.education_code({0})",
        NA_FILL_VALUE: 'secondary_or_below'
    },
    GOVERNORATE: {
        CATEGORIES: [
            'ajloun', 'al_aqaba', 'al_kirk',
            'al_mafraq', 'amman', 'balqa',
            'irbid', 'jarash', 'maadaba',
            'maan', 'tafileh', 'zarqa'
        ],
        CODE: "utils.governorate_code({0})",
        NA_FILL_VALUE: 'amman'
    },
    GENDER: {
        CATEGORIES: ['male', 'female'],
        CODE: "utils.gender_code({0})",
        NA_FILL_VALUE: 'male'
    },

}


def disability_code(value):
    if value is None:
        return None
    value = str(value).replace('.', '_').replace(' ', '_').lower()

    if value != 'no_disability' or value != '' or value is not None:
        return 'with_disability'
    return 'no_disability'


def experience_code(value):
    if value is None:
        return None
    value = float(value)
    if value > 15:
        return 20
    elif value > 10:
        return 15
    elif value > 5:
        return 10
    elif value > 1:
        return 5
    elif value > 0:
        return 1
    else:
        return 0


def age_code(value):
    if value is None:
        return None
    age = int(round(float(value) / 10) * 10)

    if 15 <= age <= 65:
        return age
    return 30


def education_code(value):
    if value is None:
        return None

    value = str(value).lower()
    education_subs = {
        'bachelor_or_above': [
            'bachelor_or_above', 'bachelor', 'bachelors', "bs", "b.s", "bas",
            'master', 'masters', 'm.s', 'ms', 'phd', 'doctor_of_philosophy', 'doctorate', 'doctorates'
        ],
        'vocational_training': [
            "vocational_training", 'vt', 'v.t'
        ],
        'middle_diploma': [
            "middle_diploma", 'diploma' "high_diploma", 'deploma', "high_deploma", "middle_deploma",
        ],
        'secondary_or_below': [
            "secondary_or_below", 'high_school', 'school', 'secondary', 'secondary_school'
        ],
    }
    for key in education_subs:
        if value in education_subs[key]:
            return key

    return TEMPLATE_DATA_CATEGORIES[EDUCATION][NA_FILL_VALUE]


def governorate_code(value):
    if value is None:
        return None

    value = str(value).lower().replace(f'governorate_', '')
    # print("Preprocessing ---------------------------------------- ", value)
    governorate_subs = {
        'al_kirk': ['al_kark', 'al_kirk', 'kark', 'kirk', ],
        'balqa': ['balqa', 'al_balqa', 'balqaa', 'al_balqaa', ],
        'tafileh': ['tafileh', 'al_tafileh', ],
        'jarash': ['jarash', 'jerash'],
        'zarqa': ['zarqa', 'al_zarqa'],
        'amman': ['amman'],
        'al_mafraq': ['al_mafraq', 'mafraq'],
        'maan': ['maan', ],
        'irbid': ['irbid', 'irbed'],
        'al_aqaba': ['al_aqaba', 'aqaba'],
        'maadaba': ['maadaba', 'madaba'],
        'ajloun': ['ajloun'],
    }
    for key in governorate_subs:
        if value in governorate_subs[key]:
            return key
    return TEMPLATE_DATA_CATEGORIES[GOVERNORATE][NA_FILL_VALUE]


def gender_code(value):
    if value is None:
        return None

    value = str(value).lower()
    gender_subs = {
        'male': ['male', 'm', 'man'],
        'female': ['female', 'f', 'woman'],
    }
    for key in gender_subs:
        if value in gender_subs[key]:
            return key
    return TEMPLATE_DATA_CATEGORIES[GENDER][NA_FILL_VALUE]


# this is a bruteforce solution. I believe there is a better solution can be found
def predict_node_for_missing_value(model, data):
    """This function build the tree structure then select a node for it iff it has missing values """
    if any(elem is None for elem in data) or any(elem is np.nan for elem in data):
        pass
    else:
        return 0, "This person doesn't have any missing values"

    n_nodes = model.tree_.node_count
    children_left = model.tree_.children_left
    children_right = model.tree_.children_right
    feature = model.tree_.feature
    name = model.feature_names_in_
    threshold = model.tree_.threshold
    samples = model.tree_.n_node_samples

    node_depth = np.zeros(shape=n_nodes, dtype=np.int64)
    is_leaves = np.zeros(shape=n_nodes, dtype=bool)
    stack = [(0, 0)]  # start with the root node id (0) and its depth (0)
    while len(stack) > 0:
        # `pop` ensures each node is only visited once
        node_id, depth = stack.pop()
        node_depth[node_id] = depth

        # If the left and right child of a node is not the same we have a split
        # node
        is_split_node = children_left[node_id] != children_right[node_id]
        # If a split node, append left and right children and depth to `stack`
        # so we can loop through them
        if is_split_node:
            stack.append((children_left[node_id], depth + 1))
            stack.append((children_right[node_id], depth + 1))
        else:
            is_leaves[node_id] = True

    print(
        "The binary tree structure has {n} nodes, {l} leaves and has the following tree structure:\n".format(n=n_nodes,
                                                                                                             l=np.sum(
                                                                                                                 is_leaves)))

    current_node = 0
    for i in range(n_nodes):
        print('current_node is', current_node, end=' ::: ')
        if is_leaves[i]:
            print("{space}node={node} is a leaf node.".format(space=node_depth[i] * "\t", node=i))
        else:
            if data[feature[i]] is np.nan or data[feature[i]] is None:
                if current_node == i:
                    print(
                        f"{node_depth[i] * '    '}node={i} is a split node: go to node {children_left[i]} if node "
                        f"{children_left[i]} counts {samples[children_left[i]]} >= node{children_right[i]} "
                        f"counts{samples[children_right[i]]} else to node {children_right[i]}."
                    )
                    if samples[children_left[i]] >= samples[children_right[i]]:
                        current_node = children_left[i]
                    elif samples[children_right[i]] > samples[children_left[i]]:
                        current_node = children_right[i]
            else:
                if current_node == i:
                    print(
                        f"{node_depth[i] * '    '}node={i} is a split node: go to node {children_left[i]} if "
                        f"X[:, {feature[i]} {name[feature[i]]}] <= {threshold[i]} else to node {children_right[i]}."
                    )
                    if data[feature[i]] <= threshold[i]:
                        current_node = children_left[i]

                    elif data[feature[i]] > threshold[i]:
                        current_node = children_right[i]

        if is_leaves[current_node]:
            print(f'Node {current_node} is a leaf node. Its the node where this person will land')
            return current_node, np.argmax(model.tree_.value[current_node])

    return current_node, np.argmax(model.tree_.value[current_node])


# def industry_code(value):
#     if value in TEMPLATE_DATA_CATEGORIES[INDUSTRY][CATEGORIES]:
#         return str(value).lower()
#
#     return TEMPLATE_DATA_CATEGORIES[INDUSTRY][NA_FILL_VALUE]


# def employment_code(value):
#     value = str(value).lower()
#     employment_subs = {
#         'unemployed': ['home', 'unemployed', 'un_employed'],
#         'formal_worker': ['job', 'worker', 'formal_worker', 'full_time'],
#         'daily_worker': ['daily_worker'],
#         'informal_worker': ['part_time', 'parttime', 'informal', 'informal_worker'],
#         'housewife': ['housewife', 'house_wife'],
#         'self_employed': ['freelancer', 'freelance', 'self_employed'],
#     }
#     for key in employment_subs:
#         if value in employment_subs[key]:
#             return key
#
#     return TEMPLATE_DATA_CATEGORIES[EMPLOYMENT][NA_FILL_VALUE]


# def first_job_code(value):
#     return 1 if value == 1 else 0
