from django.conf import settings
import os

GOVERNORATE = 'governorate'
AGE = 'age'
EXPERIENCE = 'experience'
EDUCATION = 'education'
GENDER = 'gender'

TARGET = 'clusters'



NA_FILL_VALUE = 'NA_FILL_VALUE'
CATEGORIES = 'CATEGORIES'
CODE = 'CODE'

FIRST_JOB = 'first_ind'
INDUSTRY = 'industry'
EMPLOYMENT = 'employment'
DISABILITY = 'disability'

TEMPLATE_DATA_CATEGORIES = {
    # FIRST_JOB: {  # Based on the R code. This doesn't contribute in the model.
    #     CATEGORIES: [0, 1],
    #     CODE: "first_job_code({0})",
    #     NA_FILL_VALUE: 0
    # },
    DISABILITY: {
        CATEGORIES: ['with_disability', 'no_disability'],
        CODE: "disability_code({0})",
        NA_FILL_VALUE: 'no_disability'
    },
    EMPLOYMENT: { # Based on the R code. This doesn't contribute in the model.
        CATEGORIES: ['unemployed', 'formal_worker', 'daily_worker', 'informal_worker', 'housewife', 'self_employed'],
        CODE: "employment_code({0})",
        NA_FILL_VALUE: 'unemployed'
    },
    # INDUSTRY: {  # Based on the R code. This doesn't contribute in the model.
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
    EXPERIENCE: {
        CATEGORIES: [0, 1, 5, 10, 15, 20],
        CODE: "experience_code({0})",
        NA_FILL_VALUE: 0
    },
    AGE: {
        CATEGORIES: [10, 20, 30, 40, 50, 60],
        CODE: "age_code({0})",
        NA_FILL_VALUE: 30
    },

    EDUCATION: {
        CATEGORIES: ['bachelor_or_above', 'vocational_training', 'middle_diploma', 'secondary_or_below'],
        CODE: "education_code({0})",
        NA_FILL_VALUE: 'secondary_or_below'
    },
    GOVERNORATE: {
        CATEGORIES: [
            'ajloun', 'al_aqaba', 'al_kirk',
            'al_mafraq', 'amman', 'balqa',
            'irbid', 'jarash', 'maadaba',
            'maan', 'tafileh', 'zarqa'
        ],
        CODE: "governorate_code({0})",
        NA_FILL_VALUE: 'amman'
    },
    GENDER: {  # Based on the R code. Different model for each of them
        CATEGORIES: ['male', 'female'],
        CODE: "gender_code({0})",
        NA_FILL_VALUE: 'male'
    },

}

ONE_HOT_ENCODED_FEATURES = [GOVERNORATE]
LABEL_ENCODED_FEATURES = [GENDER, EDUCATION, ]
CATEGORICAL_FEATURES = ONE_HOT_ENCODED_FEATURES + LABEL_ENCODED_FEATURES

ENCODERS = {
    GENDER: os.path.join(settings.BASE_DIR, 'PreProcessingAndModelCode', 'runs', f'{GENDER}_encoder.joblib'),
    EDUCATION: os.path.join(settings.BASE_DIR, 'PreProcessingAndModelCode', 'runs', f'{EDUCATION}_encoder.joblib'),
    GOVERNORATE: os.path.join(settings.BASE_DIR, 'PreProcessingAndModelCode', 'runs', f'{GOVERNORATE}_encoder.joblib'),
}
MODEL = os.path.join(settings.BASE_DIR, 'PreProcessingAndModelCode', 'runs', 'model.joblib')

# API_FEATURES = [
#     EXPERIENCE, AGE, EDUCATION, GENDER, 'ajloun', 'al_aqaba',
#     'al_kirk', 'al_mafraq', 'amman', 'balqa', 'irbid', 'jarash',
#     'maadaba', 'maan', 'outside_jordan', 'tafileh', 'zarqa'
# ]
API_FEATURES = [
    GENDER, EDUCATION, EXPERIENCE, AGE, GOVERNORATE
]
CONTRIBUTORS = [GENDER, EDUCATION, EXPERIENCE, AGE, GOVERNORATE, TARGET]


def disability_code(value):
    if str(value).lower() != 'no_disability' or value != '' or value is not None:
        return 'with_disability'
    return 'no_disability'


# def first_job_code(value):
#     return 1 if value == 1 else 0


def experience_code(value):
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
    age = int(round(float(value) / 10) * 10)

    if 15 <= age <= 65:
        return age
    return 30


def employment_code(value):
    value = str(value).lower()
    employment_subs = {
        'unemployed': ['home', 'unemployed', 'un_employed'],
        'formal_worker': ['job' , 'worker', 'formal_worker', 'full_time'],
        'daily_worker': ['daily_worker'],
        'informal_worker': ['part_time', 'parttime', 'informal', 'informal_worker'],
        'housewife': ['housewife', 'house_wife'],
        'self_employed': ['freelancer', 'freelance', 'self_employed'],
    }
    for key in employment_subs:
        if value in employment_subs[key]:
            return key

    return TEMPLATE_DATA_CATEGORIES[EMPLOYMENT][NA_FILL_VALUE]


def education_code(value):
    value = str(value).lower()
    print("Preprocessing ---------------------------------------- ", value)
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
    value = str(value).lower().replace(f'governorate_', '')
    print("Preprocessing ---------------------------------------- ", value)
    governorate_subs = {
        'al_kirk': ['al_kark', 'al_kirk', 'kark', 'kirk', ],
        'balqa': ['balqa', 'al_balqa', 'balqaa', 'al_balqaa',],
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
    value = str(value).lower()
    gender_subs = {
        'male': [ 'male', 'm', 'man'],
        'female': [ 'female', 'f', 'woman'],
    }
    for key in gender_subs:
        if value in gender_subs[key]:
            return key
    return TEMPLATE_DATA_CATEGORIES[GENDER][NA_FILL_VALUE]

# def industry_code(value):
#     if value in TEMPLATE_DATA_CATEGORIES[INDUSTRY][CATEGORIES]:
#         return str(value).lower()
#
#     return TEMPLATE_DATA_CATEGORIES[INDUSTRY][NA_FILL_VALUE]
