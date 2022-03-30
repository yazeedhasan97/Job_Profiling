
from mongoengine import Document, fields, DynamicDocument
from datetime import datetime
from django.conf import settings
import pandas as pd





class EvaluationOutput(DynamicDocument):
    run_time = fields.IntField()

    @classmethod
    def fill_data(cls):
        cls.df = pd.read_csv(settings.EVALUATION_FILE).fillna('')

        di = cls.to_dict('records')
        
        instances = [EvaluationOutput(**data) for data in di]
        
        EvaluationOutput.objects.insert(instances, load_bulk=False)    
        return True,len(instances)
   
