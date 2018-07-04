# -*- coding: utf-8 -*-

import logging
import logging.handlers
import urllib
import json
import requests

logger = logging.getLogger('splunk')


class JSONHttpsClient:
    def __init__(self, namespace, user, token, host, port, use_ssl, page_size=100):
        self.host = host
        self.port = port
        self.user = user
        self.token = token
        self.page_size = page_size
        self.namespace = namespace
        self.decoder = json.JSONDecoder()

        self.scheme = "http"

        if use_ssl:
            self.scheme = "https"

        self.headers = headers = {
            "User-Agent": "excel-json-client/0.1",
            "Host":  "%s:%s" % (self.host, self.port),
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "*/*",
            "Authorization": "Splunk %s" % self.token
        }
        pass

    def get(self, sid, etype, search, parameters):
        if sid:
            path = '/servicesNS/%s/%s/search/jobs/%s/%s/export'\
                % (self.user, self.namespace, sid, etype)
        else:
            path = '/servicesNS/%s/%s/search/jobs/export' % (
                self.user, self.namespace)
            if not search.strip().startswith(('search ', '|')):
                search = 'search ' + search
            parameters['search'] = search

        qstr = urllib.urlencode(parameters)
        url = self.scheme + '://' + self.host + \
            ":" + str(self.port) + path + "?" + qstr

        try:
            response = requests.get(
                url, headers=self.headers, verify=False)
        except Exception, ex:
            logger.error(
                "connection to %s on port %s failed with error %s", self.host, self.port, ex)
            raise
        try:
            result = []
            for line in response.text.splitlines():
                result.append(self.decoder.raw_decode(line))
            return result
        except Exception, ex:
            logger.error("cannot parse json: %s", response.text)
            logger.exception(ex)
            raise
