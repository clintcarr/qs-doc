import requests
from requests_ntlm import HttpNtlmAuth
import json
import csv
import random
import string


requests.packages.urllib3.disable_warnings()

def set_xrf():
    characters = string.ascii_letters + string.digits
    return ''.join(random.sample(characters, 16))

xrf = set_xrf()

headers = {"X-Qlik-XrfKey": xrf,
            "Accept": "application/json",
            "X-Qlik-User": "UserDirectory=Internal;UserID=sa_repository",
            "Content-Type": "application/json"}

session = requests.session()

class ConnectQlik:
    """
    Instantiates the Qlik Repository Service Class
    """

    def __init__(self, server, certificate = False, root = False
        ,userdirectory = False, userid = False, credential = False, password = False):
        """
        Establishes connectivity with Qlik Sense Repository Service
        :param server: servername.domain:4242
        :param certificate: path to client.pem and client_key.pem certificates
        :param root: path to root.pem certificate
        :param userdirectory: userdirectory to use for queries
        :param userid: user to use for queries
        :param credential: domain\\username for Windows Authentication
        :param password: password of windows credential
        """
        self.server = server
        self.certificate = certificate
        self.root = root
        if userdirectory is not False:
            headers["X-Qlik-User"] = "UserDirectory={0};UserID={1}".format(userdirectory, userid)
        self.credential = credential
        self.password = password

    def get(self,endpoint,filterparam=None, filtervalue=None):
        """
        Function that performs GET method to Qlik Repository Service endpoints
        :param endpoint: API endpoint path
        :param filterparam: Filter for endpoint, use None for no filtering
        :param filtervalue: Value to filter on, use None for no filtering
        """
        if self.credential is not False:
            session.auth = HttpNtlmAuth(self.credential, self.password, session)
            headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36'
        if filterparam is None:
            if '?' in endpoint:
                response = session.get('https://{0}/{1}&xrfkey={2}'.format (self.server, endpoint, xrf),
                                        headers=headers, verify=self.root, cert=self.certificate)
                return response.content
            else:
                response = session.get('https://{0}/{1}?xrfkey={2}'.format (self.server, endpoint, xrf),
                                        headers=headers, verify=self.root, cert=self.certificate)
                return response.content
        else:
            response = session.get("https://{0}/{1}?filter={2} '{3}'&xrfkey={4}".format 
                                    (self.server, endpoint, filterparam, filtervalue, xrf), 
                                    headers=headers, verify=self.root, cert=self.certificate)
            
            return response.content


    def get_qps(self, endpoint):
        """
        Function that performs GET method to Qlik Proxy Service endpoints
        :param endpoint: API endpoint path
        """
        server = self.server
        qps = server[:server.index(':')]

        response = session.get('https://{0}/{1}?xrfkey={2}'.format (qps, endpoint, xrf),
                                        headers=headers, verify=self.root, cert=self.certificate)
        return response.status_code

    def get_health(self):
        """
        Function to GET the health information from the server
        :returns: JSON data
        """
        server = self.server
        engine = server[:server.index(':')]
        engine += ':4747'
        endpoint = 'engine/healthcheck'
        response = session.get('https://{0}/{1}?xrfkey={2}'.format (engine, endpoint, xrf),
                                        headers=headers, verify=self.root, cert=self.certificate)
        return json.loads(response.text)

    def get_about(self,opt=None):
        """
        Returns system information
        :returns: JSON
        """
        path = 'qrs/about'
        return json.loads(self.get(path).decode('utf-8'))

    def get_app(self, opt=None,filterparam=None, filtervalue=None):
        """
        Returns the applications
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/app'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))

    def get_dataconnection(self, opt=None, filterparam=None, filtervalue=None):
        """
        Returns the dataconnections
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/dataconnection'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))

    def get_user(self, opt=None, filterparam=None, filtervalue=None):
        """
        Returns the users
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/user'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))   

    def get_customproperty(self, opt=None, filterparam=None, filtervalue=None):
        """
        Returns the custom properties
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/custompropertydefinition'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))

    def get_tag(self, opt= None, filterparam=None, filtervalue=None):
        """
        Returns the tags
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/tag'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))
    
    def get_task(self, opt=None, filterparam=None, filtervalue=None):
        """
        Returns the tasks
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/task'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))

    def get_systemrule(self, opt=None, filterparam=None, filtervalue=None):
        """
        Returns the system rules
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/systemrule'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))

    def get_userdirectory(self, opt=None, filterparam=None, filtervalue=None):
        """
        Returns the user directory
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/userdirectory'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))
        
    def get_extension(self, opt=None, filterparam=None, filtervalue=None):
        """
        Returns the extensions
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/extension'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))

    def get_stream(self, opt=None, filterparam=None, filtervalue=None):
        """
        Returns the streams
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/stream'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))

    def get_servernode(self, opt=None, filterparam=None, filtervalue=None):
        """
        Returns the server node
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/servernodeconfiguration'
        if opt:
            path+= '/full'
        
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))

    def get_useraccesstype(self, opt=None,filterparam=None, filtervalue=None):
        """
        Returns the users with user access type
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/license/useraccesstype'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))
    
    def get_license(self):
        """
        Returns the License
        :returns: JSON
        """
        path = 'qrs/license'
        return json.loads(self.get(path).decode('utf-8'))

    def get_loginaccesstype(self, opt=None,filterparam=None, filtervalue=None):
        """
        Returns the login access type rule
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/license/loginaccesstype'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))

    def get_appobject(self, opt=None, filterparam=None, filtervalue=None):
        """
        Returns the application objects
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/app/object'
        if opt:
            path += '/full'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))

    def get_apidescription(self, method, filterparam=None, filtervalue=None):
        """
        Returns the APIs of the QRS
        :param method: Method to return (get, put, post, delete)
        :returns: JSON
        """
        path = 'qrs/about/api/description?extended=true&method={0}&format=JSON'.format (method, filterparam, filtervalue)
        return json.loads(self.get(path).decode('utf-8'))

    def get_serverconfig(self, opt=None):
        """
        Returns the server configuration
        :returns: JSON
        """
        path = 'qrs/servernodeconfiguration'
        if opt:
            path+= '/full'
        return json.loads(self.get(path).decode('utf-8')) 

    def get_emptyserverconfigurationcontainer(self):
        """
        Returns an empty server configuration
        :returns: JSON
        """
        path = 'qrs/servernodeconfiguration/local'
        return json.loads(self.get(path).decode('utf-8')) 

    def get_contentlibrary(self, filterparam=None, filtervalue=None):
        """
        Returns the content library
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/contentlibrary'
        return json.loads(self.get(path, filterparam, filtervalue).decode('utf-8'))

    def get_appprivileges(self, appid, filterparam=None, filtervalue=None):
        """
        Returns the privileges
        :param filterparam: Property and operator of the filter
        :param filtervalue: Value of the filter
        :returns: JSON
        """
        path = 'qrs/app/{0}/privileges'.format (appid)
        return json.loads(self.get(path , filterparam, filtervalue).decode('utf-8'))

    def ping_proxy(self):
        """
        Returns status code of Proxy service
        :returns: HTTP Status Code
        """
        path = 'qps/user'
        try:
            return self.get_qps(path)
        except requests.exceptions.RequestException as exception:
            return 'Qlik Sense Proxy down'

    def get_systeminfo(self, opt=None):
        """
        Returns the system information
        :param opt: Allows the retrieval of full json response
        :returns: json response
        """
        path = 'qrs/systeminfo'
        if opt:
            path += '/full'
        return json.loads(self.get(path).decode('utf-8'))

    def get_engine(self, opt=None):
        path = 'qrs/engineservice'
        if opt:
            path += '/full'
        return json.loads(self.get(path).decode('utf-8'))

    def get_proxy(self, opt=None):
        path = 'qrs/proxyservice'
        if opt:
            path += '/full'
        return json.loads(self.get(path).decode('utf-8'))

    def get_scheduler(self, opt=None):
        path = 'qrs/schedulerservice'
        if opt:
            path += '/full'
        return json.loads(self.get(path).decode('utf-8'))
    
    def get_servicecluster(self, opt=None):
        path = 'qrs/servicecluster/full'
        return json.loads(self.get(path).decode('utf-8'))

    def get_nodeconfig(self, configid):
        """
        Returns the server configuration based on the id provided
        :returns: JSON
        """
        path = 'qrs/servernodeconfiguration/{0}'.format (configid)
        return json.loads(self.get(path).decode('utf-8'))
    
    def get_repositoryservice(self, opt=None):
        path = 'qrs/repositoryservice'
        if opt:
            path+= '/full'
        return json.loads(self.get(path).decode('utf-8'))

    def get_printingservice(self, opt=None):
        path = 'qrs/printingservice'
        if opt:
            path+= '/full'
        return json.loads(self.get(path).decode('utf-8'))

    def get_analyticconnection(self, opt=None):
        path = 'qrs/AnalyticConnection'
        if opt:
            path+= '/full'
        return json.loads(self.get(path).decode('utf-8'))

    def get_odag(self, opt=None):
        path = 'qrs/odagservice'
        if opt:
            path+= '/full'
        return json.loads(self.get(path).decode('utf-8'))

if __name__ == '__main__':
    qrs = ConnectQlik(server='qs2.qliklocal.net:4242', 
                    certificate=('C:/certs/qs2.qliklocal.net/client.pem',
                                      'C:/certs/qs2.qliklocal.net/client_key.pem'),
                    root='C:/certs/qs2.qliklocal.net/root.pem')

    qrsntlm = ConnectQlik(server='qmi-qs-sn', 
                    credential='qmi-qs-sn\\vagrant',
                    password='vagrant')
   
    serverconfig = qrsntlm.get_serverconfig(opt='full')
    roles = []
    for node in range(len(serverconfig)):
        for role in range(len(serverconfig[node]['roles'])):
            roles.append ([serverconfig[node]['hostName'], serverconfig[node]['roles'][role]['definition']])
    for item in range(len(roles)):
        if roles[item][1] == 0:
            roles[item][1] = 'SchedulerMaster'
        if roles[item][1] == 1:
            roles[item][1] = 'LicenceMaintenance'
        if roles[item][1] == 2:
            roles[item][1] = 'UserSync'
        if roles[item][1] == 3:
            roles[item][1] = 'NodeRegistration'
        if roles[item][1] == 4:
            roles[item][1] = 'AppManagement'
        if roles[item][1] == 5:
            roles[item][1] = 'CleanDatabase'

    print (roles)