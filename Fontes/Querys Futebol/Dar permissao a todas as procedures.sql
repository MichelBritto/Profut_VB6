select 'GRANT EXECUTE ON ['+name+'] TO PUBLIC '  
from sys.objects  
where type ='P' 