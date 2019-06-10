import qrspy
import json
import sys
import argparse


parser = argparse.ArgumentParser()
parser.add_argument('--server', help='Qlik Sense Server to connect to.')
parser.add_argument('--certs', help='Path to certificates.')
parser.add_argument('--user', help='Username in format domain\\user')
parser.add_argument('--wmi', help='True/False, set to True to collect WMI information from servers. This switch requires windows authentication (username and password)')
parser.add_argument('--password', help='Password of user.')
args = parser.parse_args()

if args.certs:
    if args.certs[-1] == '/':
        qrs = qrspy.ConnectQlik(server=args.server+':4242', 
                            certificate=(args.certs+'client.pem',
                                            args.certs+'client_key.pem'),
                            root=args.certs+'root.pem')
    else:
        qrs = qrspy.ConnectQlik(server=args.server+':4242', 
                            certificate=(args.certs+'/client.pem',
                                            args.certs+'/client_key.pem'),
                            root=args.certs+'/root.pem')
elif args.user:
    qrs = qrspy.ConnectQlik(server=args.server,
                            credential=args.user,
                            password=args.password)

# qrs = qrspy.ConnectQlik(server='qs2.qliklocal.net:4242', 
#                     certificate=('C:/certs/qs2.qliklocal.net/client.pem',
#                                       'C:/certs/qs2.qliklocal.net/client_key.pem'),
#                     root='C:/certs/qs2.qliklocal.net/root.pem')

def get_about():
    about = qrs.get_about()
    return about

def get_apps():
    x = qrs.get_app(opt='full')
    apps = []
    for item in range(len(x)):
        if x[item]['published'] is False:
            apps.append([x[item]['name'], x[item]['description'],'Not Published', 'Not Published',x[item]['fileSize'], x[item]['owner']['userId'], x[item]['owner']['name']])
        else:
            apps.append([x[item]['name'], x[item]['description'],x[item]['publishTime'],x[item]['stream']['name'], x[item]['fileSize'], x[item]['owner']['userId'], x[item]['owner']['name']])

    return apps

def get_streams():
    streams = qrs.get_stream()
    return streams

def get_dataconnections():
    data_connection = qrs.get_dataconnection()
    connections = []
    for item in range(len(data_connection)):
        connections.append([data_connection[item]['name'],data_connection[item]['connectionstring'],data_connection[item]['type'] ])
    return connections

def get_users():
    users = qrs.get_user(opt='full')
    user_list = []
    for item in range(len(users)):
        user_list.append([[users][0][item]['userId'],
            [users][0][item]['userDirectory'],
            [users][0][item]['name'],
            [users][0][item]['roles'],
            [users][0][item]['inactive'], 
            [users][0][item]['removedExternally'],
            [users][0][item]['blacklisted']])

    return user_list

def get_systemrules(ruletype):
    system_rules = qrs.get_systemrule(filterparam='type eq', filtervalue=ruletype)
    rules = []
    for item in range(len(system_rules)):
        rules.append([system_rules[item]['name'],
                    system_rules[item]['rule'],
                    system_rules[item]['resourceFilter'],
                    system_rules[item]['actions'],
                    system_rules[item]['disabled']])

    return rules

def get_licenserules():
    system_rules = qrs.get_systemrule(filterparam='category eq', filtervalue='License')
    rules = []
    for item in range(len(system_rules)):
        rules.append([system_rules[item]['name'],
                    system_rules[item]['rule'],
                    system_rules[item]['resourceFilter'],
                    system_rules[item]['actions'],
                    system_rules[item]['disabled']])

    return rules

def get_userdirectory():
    udc = qrs.get_userdirectory(opt='full')
    directory = []
    for item in range(len(udc)):
        directory.append ([udc[item]['name'],
                    udc[item]['userDirectoryName'],
                    udc[item]['configured'],
                    udc[item]['operational'],
                    udc[item]['type'],
                    udc[item]['syncOnlyLoggedInUsers']])

    return directory

def get_customprop():
    customprop = qrs.get_customproperty(opt='full')
    properties = []
    for item in range(len(customprop)):
        properties.append ([customprop[item]['name'],
                customprop[item]['choiceValues'],
                customprop[item]['objectTypes']]
                )
    return properties

def get_tag():
    tag = qrs.get_tag()
    tags = []
    for item in range(len(tag)):
        tags.append (tag[item]['name'])
    return tags

def get_extensions():
    extension = qrs.get_extension()
    extensions = []
    for item in range(len(extension)):
        extensions.append(extension[item]['name'])
    return extensions

def get_servernode():
    node = qrs.get_servernode(opt='full')
    nodes = []
    for item in range(len(node)):
        nodes.append ([node[item]['name'], node[item]['hostName'], node[item]['serviceCluster']['name'], node[item]['failoverCandidate']])
    return nodes

def get_engine():
    engines = qrs.get_engine(opt='full')
    engine = []
    for item in range(len(engines)):
        engine.append ([engines[item]['customProperties'],
                engines[item]['settings']['listenerPorts'],
                engines[item]['settings']['autosaveInterval'],
                engines[item]['settings']['tableFilesDirectory'],
                engines[item]['settings']['genericUndoBufferMaxSize'],
                engines[item]['settings']['documentTimeout'],
                engines[item]['settings']['documentDirectory'],
                engines[item]['settings']['allowDataLineage'],
                engines[item]['settings']['qrsHttpNotificationPort'],
                engines[item]['settings']['standardReload'],
                engines[item]['settings']['workingSetSizeLoPct'],
                engines[item]['settings']['workingSetSizeHiPct'],
                engines[item]['settings']['workingSetSizeMode'],
                engines[item]['settings']['cpuThrottlePercentage'],
                engines[item]['settings']['maxCoreMaskPersisted'],
                engines[item]['settings']['maxCoreMask'],
                engines[item]['settings']['maxCoreMaskHiPersisted'],
                engines[item]['settings']['maxCoreMaskHi'],
                engines[item]['settings']['objectTimeLimitSec'],
                engines[item]['settings']['exportTimeLimitSec'],
                engines[item]['settings']['reloadTimeLimitSec'],
                engines[item]['settings']['hyperCubeMemoryLimit'],
                engines[item]['settings']['exportMemoryLimit'],
                engines[item]['settings']['reloadMemoryLimit'],
                engines[item]['settings']['createSearchIndexOnReloadEnabled'],
                engines[item]['serverNodeConfiguration']['hostName'],

                engines[item]['settings']['globalLogMinuteInterval'],
                engines[item]['settings']['auditActivityLogVerbosity'],  # ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Basic', '5: Extended'],
                engines[item]['settings']['auditSecurityLogVerbosity'],  # ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Basic', '5: Extended'],
                engines[item]['settings']['serviceLogVerbosity'], #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                engines[item]['settings']['systemLogVerbosity'],   #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                engines[item]['settings']['performanceLogVerbosity'],  #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                engines[item]['settings']['qixPerformanceLogVerbosity'], #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                engines[item]['settings']['auditLogVerbosity'],  #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                engines[item]['settings']['sessionLogVerbosity'],  #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                engines[item]['settings']['trafficLogVerbosity'],  #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'

                engines[item]['serverNodeConfiguration']['name']
                ])
    for item in range(len(engines)):
        for log in range(27, 29):
            if engine[item][log] == 0:
                engine[item][log] = 'Off'
            elif engine[item][log] == 1:
                engine[item][log] = 'Fatal'
            elif engine[item][log] == 2:
                engine[item][log] = 'Error'
            elif engine[item][log] == 3:
                engine[item][log] = 'Warning'
            elif engine[item][log] == 4:
                engine[item][log] = 'Basic'
            elif engine[item][log] == 5:
                engine[item][log] = 'Extended'

        for log in range(29, 36):
            if engine[item][log] == 0:
                engine[item][log] = 'Off'
            elif engine[item][log] == 1:
                engine[item][log] = 'Fatal'
            elif engine[item][log] == 2:
                engine[item][log] = 'Error'
            elif engine[item][log] == 3:
                engine[item][log] = 'warning'
            elif engine[item][log] == 4:
                engine[item][log] = 'Info'
            elif engine[item][log] == 5:
                engine[item][log] = 'Debug'               
        if engine[item][12]==0:
            engine[item][12] = 'Ignore Max Limit'
        elif engine[item][12] == 1:
            engine[item][12] = 'Soft Max Limit'
        elif engine[item][12] == 2:
            engine[item][12] = 'Hard Max Limit'
    return engine

def get_scheduler():
    schedulers = qrs.get_scheduler(opt='full')
    scheduler = []
    for item in range(len(schedulers)):
        scheduler.append([schedulers[item]['customProperties'],
                        schedulers[item]['settings']['schedulerServiceType'],
                        schedulers[item]['settings']['maxConcurrentEngines'],
                        schedulers[item]['settings']['engineTimeout'],
                        schedulers[item]['tags'],
                        schedulers[item]['serverNodeConfiguration']['hostName'],
                        schedulers[item]['settings']['logVerbosity']['logVerbosityAuditActivity'],  # ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Basic', '5: Extended'],
                        schedulers[item]['settings']['logVerbosity']['logVerbosityAuditSecurity'], # ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Basic', '5: Extended'],
                        schedulers[item]['settings']['logVerbosity']['logVerbosityService'], #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                        schedulers[item]['settings']['logVerbosity']['logVerbosityApplication'], #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                        schedulers[item]['settings']['logVerbosity']['logVerbosityAudit'], #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                        schedulers[item]['settings']['logVerbosity']['logVerbosityPerformance'],#['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                        schedulers[item]['settings']['logVerbosity']['logVerbositySecurity'], #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                        schedulers[item]['settings']['logVerbosity']['logVerbositySystem'], #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                        schedulers[item]['settings']['logVerbosity']['logVerbosityTaskExecution'],#['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                        schedulers[item]['serverNodeConfiguration']['name']
                        ])
    for sch in range(len(scheduler)):
            if scheduler[sch][1] == 0:
                scheduler[sch][1] = 'Master'
            elif scheduler[sch][1] == 1:
                scheduler[sch][1] = 'Slave'
            elif scheduler[sch][1] == 2:
                scheduler[sch][1] = 'Master and Slave'

    for item in range(len(scheduler)):
        for log in range(6, 8):
            if scheduler[item][log] == 0:
                scheduler[item][log] = 'Off'
            elif scheduler[item][log] == 1:
                scheduler[item][log] = 'Fatal'
            elif scheduler[item][log] == 2:
                scheduler[item][log] = 'Error'
            elif scheduler[item][log] == 3:
                scheduler[item][log] = 'Warning'
            elif scheduler[item][log] == 4:
                scheduler[item][log] = 'Basic'
            elif scheduler[item][log] == 5:
                scheduler[item][log] = 'Extended'

        for log in range(8, 15):
            if scheduler[item][log] == 0:
                scheduler[item][log] = 'Off'
            elif scheduler[item][log] == 1:
                scheduler[item][log] = 'Fatal'
            elif scheduler[item][log] == 2:
                scheduler[item][log] = 'Error'
            elif scheduler[item][log] == 3:
                scheduler[item][log] = 'warning'
            elif scheduler[item][log] == 4:
                scheduler[item][log] = 'Info'
            elif scheduler[item][log] == 5:
                scheduler[item][log] = 'Debug'    
    return scheduler

def get_proxy():
    proxies = qrs.get_proxy(opt='full')
    proxy = []
    for item in range(len(proxies)):
        proxy.append([proxies[item]['customProperties'],
                    proxies[item]['settings']['listenPort'],
                    proxies[item]['settings']['restListenPort'],
                    proxies[item]['settings']['allowHttp'],
                    proxies[item]['settings']['unencryptedListenPort'],
                    proxies[item]['settings']['authenticationListenPort'],
                    proxies[item]['settings']['kerberosAuthentication'],
                    proxies[item]['settings']['unencryptedAuthenticationListenPort'],
                    proxies[item]['settings']['keepAliveTimeoutSeconds'],
                    proxies[item]['settings']['maxHeaderSizeBytes'],
                    proxies[item]['settings']['maxHeaderLines'],
                    proxies[item]['serverNodeConfiguration']['hostName'],
                    proxies[item]['settings']['logVerbosity']['logVerbosityAuditActivity'], # ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Basic', '5: Extended'],
                    proxies[item]['settings']['logVerbosity']['logVerbosityAuditSecurity'], # ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Basic', '5: Extended'],
                    proxies[item]['settings']['logVerbosity']['logVerbosityService'], #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                    proxies[item]['settings']['logVerbosity']['logVerbosityAudit'],#['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                    proxies[item]['settings']['logVerbosity']['logVerbosityPerformance'],#['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                    proxies[item]['settings']['logVerbosity']['logVerbositySecurity'],#['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                    proxies[item]['settings']['logVerbosity']['logVerbositySystem'],#['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
                    proxies[item]['settings']['performanceLoggingInterval'],
                    proxies[item]['serverNodeConfiguration']['name']])
    
    for item in range(len(proxy)):
            if proxy[item][3] == 0:
                proxy[item][3] = False
            else:
                proxy[item][3] = True
            if proxy[item][6] == 0:
                proxy[item][6] = False
            else:
                proxy[item][6] = True

    for item in range(len(proxy)):
        for log in range(12, 14):
            if proxy[item][log] == 0:
                proxy[item][log] = 'Off'
            elif proxy[item][log] == 1:
                proxy[item][log] = 'Fatal'
            elif proxy[item][log] == 2:
                proxy[item][log] = 'Error'
            elif proxy[item][log] == 3:
                proxy[item][log] = 'Warning'
            elif proxy[item][log] == 4:
                proxy[item][log] = 'Basic'
            elif proxy[item][log] == 5:
                proxy[item][log] = 'Extended'

        for log in range(14, 19):
            if proxy[item][log] == 0:
                proxy[item][log] = 'Off'
            elif proxy[item][log] == 1:
                proxy[item][log] = 'Fatal'
            elif proxy[item][log] == 2:
                proxy[item][log] = 'Error'
            elif proxy[item][log] == 3:
                proxy[item][log] = 'warning'
            elif proxy[item][log] == 4:
                proxy[item][log] = 'Info'
            elif proxy[item][log] == 5:
                proxy[item][log] = 'Debug'    
    return proxy

def get_virtualproxy():
    virtualproxies = qrs.get_proxy(opt='full')
    alist = []
    for node in range(len(virtualproxies)):
        for vp in range(len(virtualproxies[node]['settings']['virtualProxies'])):
            alist.append([virtualproxies[node]['settings']['virtualProxies'][vp]['description'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['prefix'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['authenticationModuleRedirectUri'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['sessionModuleBaseUri'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['loadBalancingModuleBaseUri'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['authenticationMethod'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['headerAuthenticationMode'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['headerAuthenticationHeaderName'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['headerAuthenticationStaticUserDirectory'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['headerAuthenticationDynamicUserDirectory'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['anonymousAccessMode'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['windowsAuthenticationEnabledDevicePattern'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['sessionCookieHeaderName'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['sessionCookieDomain'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['additionalResponseHeaders'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['sessionInactivityTimeout'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['extendedSecurityEnvironment'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['websocketCrossOriginWhiteList'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['defaultVirtualProxy'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['tags'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['samlMetadataIdP'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['samlHostUri'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['samlEntityId'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['samlAttributeUserId'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['samlAttributeUserDirectory'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['samlAttributeSigningAlgorithm'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['samlAttributeMap'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['magicLinkHostUri'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['magicLinkFriendlyName'],

                        virtualproxies[node]['settings']['virtualProxies'][vp]['jwtPublicKeyCertificate'],    
                        virtualproxies[node]['settings']['virtualProxies'][vp]['jwtAttributeUserDirectory'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['jwtAttributeMap'],
                        virtualproxies[node]['settings']['virtualProxies'][vp]['jwtAttributeUserId'],

                        virtualproxies[node]['serverNodeConfiguration']['name']
                        ])
 
    


        for item in range(len(alist)):
            if alist[item][5] == 0:
                alist[item][5] = 'Ticket'
            elif alist[item][5] == 1:
                alist[item][5] = 'HeaderStaticUserDirectory'
            elif alist[item][5] == 2:
                alist[item][5] = 'HeaderDynamicUserDirectory'
            elif alist[item][5] == 3:
                alist[item][5] = 'SAML'
            elif alist[item][5] == 4:
                alist[item][5] = 'JWT'
            if alist[item][6] == 0:
                alist[item][6] = 'NotAllowed'
            elif alist[item][6] == 1:
                alist[item][6] = 'StaticUserDirectory'
            elif alist[item][6] == 2:
                alist[item][6] = 'DynamicUserDirectory'
            elif alist[item][6] == 3:
                alist[item][6] = 'Undefined'
            if alist[item][11] == 0:
                alist[item][11] = 'No Anonymous user'
            elif alist[item][11] == 1:
                alist[item][11] = 'Allow Anonymous'
            elif alist[item][11] == 2:
                alist[item][11] = 'Always Anonymous'
            if alist[item][26] == 0:
                alist[item][26] = 'SHA1'
            elif alist[item][26] == 1:
                alist[item][26] = 'SHA256'

    return alist

def get_vploadbalancers():
    virtualproxies = qrs.get_proxy(opt='full')
    alist = []
    for node in range(len(virtualproxies)):
        for vp in range(len(virtualproxies[node]['settings']['virtualProxies'])):
            for lb in range(len(virtualproxies[node]['settings']['virtualProxies'][vp]['loadBalancingServerNodes'])):
                alist.append([virtualproxies[node]['settings']['virtualProxies'][vp]['loadBalancingServerNodes'][lb]['hostName'],
                            virtualproxies[node]['settings']['virtualProxies'][vp]['loadBalancingServerNodes'][lb]['name'],
                            virtualproxies[node]['settings']['virtualProxies'][vp]['description']
                            ])
    sortkeyfn = key=lambda s:s[2]
    input = alist
    input.sort(key=sortkeyfn)

    from itertools import groupby
    result = []
    for key, valuesiter in groupby(input, key=sortkeyfn):
        result.append(dict(type=key, items=list(v[0]for v in valuesiter)))
    aresult = {}
    for key, valuesiter in groupby(input, key=sortkeyfn):
        aresult[key] = list(v[0] for v in valuesiter)
    return result

def get_associatedproxies():
    virtualproxies = qrs.get_proxy(opt='full')
    alist = []
    for node in range(len(virtualproxies)):
        for vp in range(len(virtualproxies[node]['settings']['virtualProxies'])):
                alist.append([virtualproxies[node]['settings']['virtualProxies'][vp]['description'],
                            virtualproxies[node]['serverNodeConfiguration']['name']
                            ])
    sortkeyfn = key=lambda s:s[1]
    input = alist
    input.sort(key=sortkeyfn)

    from itertools import groupby
    result = []
    for key, valuesiter in groupby(input, key=sortkeyfn):
        result.append(dict(type=key, items=list(v[0]for v in valuesiter)))
    aresult = {}
    for key, valuesiter in groupby(input, key=sortkeyfn):
        aresult[key] = list(v[0] for v in valuesiter)
    return  aresult

def get_udcdetails():
    udc = qrs.get_userdirectory(opt='full')
    directoryconnectors = []
    for node in range(len(udc)):
        directoryconnectors.append (udc[node]['name'])
    udc_settings=  {}

    for node in range(len(directoryconnectors)):
        udc_settings[directoryconnectors[node]] = ({'setting0': udc[node]['settings'][0]['value']})
        for setting in range(len(udc[node]['settings'])):         
            if len(udc[node]['settings']) > 1:
                udc_settings[directoryconnectors[node]].update ({'setting{0}'.format(setting): udc[node]['settings'][setting]['value']})

    return udc_settings
    
def get_tasks():
    tasks = qrs.get_task(opt='full')
    task_list = []
    for task in range(len(tasks)):
        if tasks[task]['taskType'] == 2:
            task_list.append ([tasks[task]['name'],
                        tasks[task]['userDirectory']['name'],
                        tasks[task]['enabled'],
                        tasks[task]['customProperties'],
                        tasks[task]['tags'],
                        tasks[task]['schemaPath']] )
        elif tasks[task]['taskType'] == 0:
            task_list.append ([tasks[task]['name'],
                        tasks[task]['app']['name'],
                        tasks[task]['enabled'],
                        tasks[task]['customProperties'],
                        tasks[task]['tags'],
                        tasks[task]['schemaPath']])
    return task_list

def get_license():
    qs_license = qrs.get_license()
    lic = []
    lic.append([[qs_license['lef']], [qs_license['serial']], [qs_license['name']], [qs_license['organization']], [qs_license['product']],
            [qs_license['numberOfCores']], [qs_license['isExpired']], [qs_license['expiredReason']],[qs_license['isBlacklisted']]] )
    return lic[0]

def get_servicecluster():
    service_cluster = qrs.get_servicecluster()
    sc = []
    sc.append([ service_cluster[0]['name']
                ,service_cluster[0]['settings']['sharedPersistenceProperties']['rootFolder']
                , service_cluster[0]['settings']['sharedPersistenceProperties']['appFolder']
                , service_cluster[0]['settings']['sharedPersistenceProperties']['staticContentRootFolder']
                , service_cluster[0]['settings']['sharedPersistenceProperties']['connector32RootFolder']
                , service_cluster[0]['settings']['sharedPersistenceProperties']['connector64RootFolder']
                , service_cluster[0]['settings']['sharedPersistenceProperties']['archivedLogsRootFolder']
                , service_cluster[0]['settings']['sharedPersistenceProperties']['databaseHost']
                , service_cluster[0]['settings']['sharedPersistenceProperties']['databasePort']
                , service_cluster[0]['settings']['sharedPersistenceProperties']['failoverTimeout']
                 ])
    return sc[0]

def get_nodeconfig():
    servernodeconfig = qrs.get_serverconfig(opt='full')
    nodes = []
    nodeconfiguration = []
    for node in range(len(servernodeconfig)):
        nodeconfiguration.append  (qrs.get_nodeconfig(servernodeconfig[node]['id'])
                )
    node_output = []
    for node in range(len(nodeconfiguration)):
        node_output.append ([
                nodeconfiguration[node]['hostName'],
                nodeconfiguration[node]['isCentral'],
                nodeconfiguration[node]['nodePurpose'],
                nodeconfiguration[node]['engineEnabled'],
                nodeconfiguration[node]['proxyEnabled'],
                nodeconfiguration[node]['printingEnabled'],
                nodeconfiguration[node]['schedulerEnabled'],
                nodeconfiguration[node]['name']
        ])

    for item in range(len(node_output)):
        if node_output[item][2] == 0:
            node_output[item][2] = 'Production'
        elif node_output[item][2] == 1:
            node_output[item][2] = 'Development'
        else:
            node_output[item][2] = 'Production and Development'
    return node_output

def get_logverbosity():
    logverbosity = qrs.get_repositoryservice(opt='full')
    repologverbosity = [] 
    for node in range(len(logverbosity)):
        repologverbosity.append ([
              logverbosity[node]['settings']['logVerbosity']['logVerbosityAuditActivity'], #{'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Basic', '5: Extended']
              logverbosity[node]['settings']['logVerbosity']['logVerbosityAuditSecurity'], #{'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Basic', '5: Extended']
              logverbosity[node]['settings']['logVerbosity']['logVerbosityService'], #{'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug']
              logverbosity[node]['settings']['logVerbosity']['logVerbosityApplication'], #{'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug']
              logverbosity[node]['settings']['logVerbosity']['logVerbosityAudit'], #{'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug']
              logverbosity[node]['settings']['logVerbosity']['logVerbosityLicense'], #{'values': ['0: Info', '1: Debug'],
              logverbosity[node]['settings']['logVerbosity']['logVerbosityRuleAudit'], # {'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug']
              logverbosity[node]['settings']['logVerbosity']['logVerbosityManagementConsole'], #{'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug']
              logverbosity[node]['settings']['logVerbosity']['logVerbosityPerformance'],  #{'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug']
              logverbosity[node]['settings']['logVerbosity']['logVerbositySecurity'], #{'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug']
              logverbosity[node]['settings']['logVerbosity']['logVerbositySynchronization'], #{'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug']
              logverbosity[node]['settings']['logVerbosity']['logVerbositySystem'], #{'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug']
              logverbosity[node]['settings']['logVerbosity']['logVerbosityUserManagement'],  # {'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug']
              logverbosity[node]['serverNodeConfiguration']['name']]
              )
    
    for item in range(len(repologverbosity)):
        if repologverbosity[item][5] == 0:
            repologverbosity[item][5] = 'Info'
        elif repologverbosity[item][5] == 1:
            repologverbosity[item][5] = 'Debug'
        
        if repologverbosity[item][0] == 0:
            repologverbosity[item][0] = 'Off'
        elif repologverbosity[item][0] == 1:
            repologverbosity[item][0] = 'Fatal'
        elif repologverbosity[item][0] == 2:
            repologverbosity[item][0] = 'Error'
        elif repologverbosity[item][0] == 3:
            repologverbosity[item][0] = 'Warning'
        elif repologverbosity[item][0] == 4:
            repologverbosity[item][0] = 'Basic'
        elif repologverbosity[item][0] == 5:
            repologverbosity[item][0] = 'Extended'
        
        if repologverbosity[item][1] == 0:
            repologverbosity[item][1] = 'Off'
        elif repologverbosity[item][1] == 1:
            repologverbosity[item][1] = 'Fatal'
        elif repologverbosity[item][1] == 2:
            repologverbosity[item][1] = 'Error'
        elif repologverbosity[item][1] == 3:
            repologverbosity[item][1] = 'Warning'
        elif repologverbosity[item][1] == 4:
            repologverbosity[item][1] = 'Basic'
        elif repologverbosity[item][1] == 5:
            repologverbosity[item][1] = 'Extended'
        
        for log in range(len(repologverbosity[item])):
            if repologverbosity[item][log] == 0:
                repologverbosity[item][log] = 'Off'
            elif repologverbosity[item][log] == 1:
                repologverbosity[item][log] = 'Fatal'
            elif repologverbosity[item][log] == 2:
                repologverbosity[item][log] = 'Error'
            elif repologverbosity[item][log] == 3:
                repologverbosity[item][log] = 'Warning'
            elif repologverbosity[item][log] == 4:
                repologverbosity[item][log] = 'Info'
            elif repologverbosity[item][log] == 5:
                repologverbosity[item][log] = 'Debug'
    return repologverbosity

def get_printing():
    printing = qrs.get_printingservice(opt='full')
    print_svc = []
    for node in range(len(printing)):
        print_svc.append ([printing[node]['customProperties'],
    printing[node]['settings']['workingSetSizeHiPct'],
    printing[node]['settings']['logVerbosity']['logVerbosityAuditActivity'], #['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Basic', '5: Extended'],
    printing[node]['settings']['logVerbosity']['logVerbosityService'], #'values': ['0: Off', '1: Fatal', '2: Error', '3: Warning', '4: Info', '5: Debug'
    printing[node]['serverNodeConfiguration']['hostName'],
    printing[node]['tags'],
    printing[node]['serverNodeConfiguration']['name']
        ])
    for item in range(len(print_svc)):
        if print_svc[item][2]==0:
            print_svc[item][2] = 'Off'
        elif print_svc[item][2] == 1:
            print_svc[item][2] = 'Fatal'
        elif print_svc[item][2] == 2:
            print_svc[item][2] = 'Error'
        elif print_svc[item][2] == 3:
            print_svc[item][2] = 'Warning'
        elif print_svc[item][2] == 4:
           print_svc[item][2] = 'Basic'
        elif print_svc[item][2] == 5:
            print_svc[item][2] = 'Extended'
    
    for item in range(len(print_svc)):
        if print_svc[item][3]==0:
            print_svc[item][3] = 'Off'
        elif print_svc[item][3] == 1:
            print_svc[item][3] = 'Fatal'
        elif print_svc[item][3] == 2:
            print_svc[item][3] = 'Error'
        elif print_svc[item][3] == 3:
            print_svc[item][3] = 'Warning'
        elif print_svc[item][3] == 4:
           print_svc[item][3] = 'Info'
        elif print_svc[item][3] == 5:
            print_svc[item][3] = 'Debug'
    return print_svc

def get_analyticconnections():
    analytics = qrs.get_analyticconnection(opt='full')
    aconnections = []
    for conn in range(len(analytics)):
        aconnections.append ([analytics[conn]['host'],
                        analytics[conn]['port'],
                        analytics[conn]['certificateFilePath'],
                        analytics[conn]['reconnectTimeout'],
                        analytics[conn]['requestTimeout'],
                        analytics[conn]['name']])
    
    return aconnections

def get_odag():
    odag = qrs.get_odag(opt='full')
    odagconnections = []
    odagconnections.append ([odag[0]['settings']['enabled'],
                        odag[0]['settings']['maxConcurrentRequests'],
                        odag[0]['settings']['logLevel']])
    for item in range(len(odagconnections)):
        if odagconnections[item][2] == 0:
            odagconnections[item][2] = 'Off'
        if odagconnections[item][2] == 1:
            odagconnections[item][2] = 'Fatal'
        if odagconnections[item][2] == 2:
            odagconnections[item][2] = 'Error'
        if odagconnections[item][2] == 3:
            odagconnections[item][2] = 'Warning'
        if odagconnections[item][2] == 4:
            odagconnections[item][2] = 'Info'
        if odagconnections[item][2] == 5:
            odagconnections[item][2] = 'Debug'
        if odagconnections[item][2] == 6:
            odagconnections[item][2] = 'Trace'
    return odagconnections

def masterRoles():
    serverconfig = qrs.get_serverconfig(opt='full')
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
    return roles

if __name__ == '__main__':
    qrs = ConnectQlik(server='qmi-qs-sn', 
                    credential='qmi-qs-sn\\vagrant',
                    password='vagrant')
