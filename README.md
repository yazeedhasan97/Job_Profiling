# job_profiling # Under development

An API to run a machine learning model that classify the people into clusters based on some of there living status

Available clusters:

	Cluster1 : Has no problem finding a Job
	
	Cluster2 : Can find a job, yet needs some consideration
	
	Cluster3 : In crutial need for training to find a job

end points : returns the cluster based on provided info in the [GET] request :

Notes:

	1- None of the attributes is requaired (A placeholder will take a part)
	
	
	2- Missing Attributes shouldn't be passsed. However, wrong passed values will be treated as Null
	
	
	3- Additions attributes can be passed. However, they will be execluded (This is to make it more dynamic for future attributes to contribute)
	
	

inputs are 

	1- 'experience' --> [any float between 0 and 150]
	
	
	2- 'age' --> [any float between 0 and 150]
	
	
	3- 'education' --> ['bachelor_or_above', 'vocational_training', 'middle_diploma', 'secondary_or_below']	
	
	
	4- 'governorate' --> ['ajloun', 'al_aqaba', 'al_kirk', 'al_mafraq', 'amman', 'balqa', 'irbid', 'jarash', 'maadaba', 'maan', 'tafileh', 'zarqa']
	
	
	5- 'gender' --> ['male', 'female']
	
	6- 'disability' --> ['no_disability', 'with_disability']

http://127.0.0.1:8000/cluster/?experience=0&age=57&education=secondary_or_below&governorate=al_kirk&gender=female&disability=no_disability # this is how full&right api call should look


http://127.0.0.1:8000/cluster/?experience=0&age=28&education=middle_diploma&governorate=ajloun&gender=female


http://127.0.0.1:8000/cluster/?experience=0&age=28.6&education=middle_diploma&governorate=ajloun&gender=female


http://127.0.0.1:8000/cluster/?experience=4.5&age=32&education=secondary_or_below&governorate=maan&gender=female


http://127.0.0.1:8000/cluster/?experience=0&age=23&education=vocational_training&governorate=maan&gender=female


http://127.0.0.1:8000/cluster/?experience=3.2&age=22&education=secondary_or_below&governorate=tafileh&gender=female


http://127.0.0.1:8000/cluster/?experience=0&age=24.2&education=vocational_training&governorate=Aqaba&gender=female


http://127.0.0.1:8000/cluster/?experience=0&governorate=ِAmman&education=bachelor_or_above&age=27.32&gender=female


http://127.0.0.1:8000/cluster/?experience=0&governorate=ِAmman&education=bachelor_or_above&age=21&gender=female


http://127.0.0.1:8000/cluster/?experience=1&governorate=ِAmman&education=phd&age=29&gender=female


http://127.0.0.1:8000/cluster/?experience=5&governorate=ِAmman&education=bachelor_or_above&age=29&gender=female


http://127.0.0.1:8000/cluster/?experience=5&governorate=ِAmman&education=bachelor_or_above&age=29&gender=female&disability=no_disability # you can provide additional variables but they will be excluded


http://127.0.0.1:8000/cluster/?experience=5&governorate=ِAmman&education=bachelor_or_above&age=29&gender=female&disability=no_disability&employment=house_wife # you can provide additional variables but they will be excluded


http://127.0.0.1:8000/cluster/?experience=4&governorate=ِAmman&education=deploma&age=22.5&gender=male


http://127.0.0.1:8000/cluster/?experience=1&governorate=irbid&education=bachelor&age=24&gender=male


http://127.0.0.1:8000/cluster/?experience=0&governorate=irbid&education=bachelor&age=24&gender=male


http://127.0.0.1:8000/cluster/?experience=0&governorate=al_balqa&education=school&age=19&gender=male


http://127.0.0.1:8000/cluster/?experience=1&governorate=amman&education=school&age=19&gender=male


http://127.0.0.1:8000/cluster/?experience=5&governorate=zarqa&education=vocational_training&age=19&gender=male


http://127.0.0.1:8000/cluster/?experience=5&governorate=zarqa&education=vocational_training&age=19 # missing attributes shouldn't be provided


http://127.0.0.1:8000/cluster/?experience=5&governorate=zarqa&education=vocational_training&age=19&gender=null # attributes provided with wrong values will be considared as nulls














