from django.urls import path, include
from django.contrib.auth.models import User
from rest_framework import routers, serializers, viewsets
from api.documents import EvaluationOutput

from rest_framework_mongoengine.serializers import DynamicDocumentSerializer


# class EvaluationOutputSerializer(serializers.ModelSerializer):
#     class Meta:
#         model = EvaluationOutput
#         fields = '__all__'

class EvaluationOutputSerializer(DynamicDocumentSerializer):
    
    class Meta:
        model = EvaluationOutput
        fields = '__all__'



