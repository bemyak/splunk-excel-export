import cherrypy
import splunk.entity as entity
import splunk.util as util

def host_app(fn):
    '''
    usage:
    
    @host_app
    def my_func(self, host_app=None):
        assert(isinstance(host_app, basestr))
    '''
    def decorator(self, *args, **kwargs):
        kwargs.update({'host_app' : cherrypy.request.path_info.split('/')[3]})
        return fn(self, *args, **kwargs)
    
    return decorator
    
def version_info(fn):
    '''
    usage:
    
    @version_info
    def my_func(self, build=None, isFree=False, isTrial=True, version=None):
        assert(isinstance(build, int))
        assert(isinstance(isFree, bool))
        assert(isinstance(isTrial, bool))
        assert(isinstance(version, basestr))
    '''
    def decorator(self, *args, **kwargs):
        sessionKey = cherrypy.session.get('sessionKey')
        en = entity.getEntity('services/server/info', 'server-info',
                              sessionKey=sessionKey)
        kwargs.update({'build': int(en['build']),
                       'isFree': util.normalizeBoolean(en['isFree']),
                       'isTrial': util.normalizeBoolean(en['isTrial']),
                       'version': en['version']})
        return fn(self, *args, **kwargs)

    return decorator

