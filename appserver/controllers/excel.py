# -*- coding: utf-8 -*-

import logging
import os

import cherrypy

import splunk
import splunk.appserver.mrsparkle.controllers as controllers
from splunk.appserver.mrsparkle.lib.decorators import expose_page
from splunk.appserver.mrsparkle.lib.routes import route
import splunk.rest
import splunk.search
import splunk.util

from excel_export.clients.json_https_client import JSONHttpsClient
from excel_export.clients.wb_client import WBClient
from excel_export.decorators import host_app

logger = logging.getLogger('splunk')

SPLUNK_HOME = os.environ.get('SPLUNK_HOME')


class excel(controllers.BaseController):
    """Excel Export Controller"""

    @route('/')
    @expose_page(must_login=True, methods=['GET'])
    @host_app
    def default(self, host_app=None, filename=None, template=None, count=10000, **kwargs):
        """ create excel report """
        try:
            count = abs(int(count))
        except:
            logger.warn(
                'given count %s is not a positive integer, reverting to 10000' % count)
            count = 10000

        if count > 1000000:
            count = 1000000
            logger.warn(
                'count %s was reduced so as not to exceed excel max row count' % count)

        if not filename:
            filename = 'splunk_report.xlsx'
        elif not filename.endswith('.xlsx'):
            logger.warn(
                'xls file extension will be appended to given filename %s ' % filename)
            filename = '.'.join([filename, 'xlsx'])

        user = cherrypy.session['user']['name']
        token = cherrypy.session.get('sessionKey')
        host = splunk.getDefault('host')
        port = splunk.getDefault('port')
        use_ssl = splunk.util.normalizeBoolean(splunk.rest.simpleRequest(
            '/services/properties/server/sslConfig/enableSplunkdSSL')[1])

        try:
            fetcher = JSONHttpsClient(
                host_app, user, token, host, port, use_ssl)
        except Exception, ex:
            logger.exception(ex)
            raise

        cherrypy.response.stream = True
        cherrypy.response.headers['Content-Type'] = 'application/mx-excel'
        cherrypy.response.headers['Content-Disposition'] = 'attachment; filename="%s"' % filename

        try:
            wb_client = WBClient(template=template)
        except Exception, ex:
            logger.exception(ex)
            raise

        for key in sorted(kwargs):
            logger.warn("key=%s; value=%s", key, kwargs[key])
            sid = kwargs[key]
            use_sid = None
            etype = None
            logger.info(
                'creating excel export for user %s from search id %s' % (user, sid))

            try:
                job = self.get_processed_job(sid)
            except Exception, ex:
                logger.exception(ex)
                break

            results = getattr(job, 'results')
            field_names = [x for x in results.fieldOrder
                           if (not x.startswith('_') or x == '_time' or x == '_raw')]
            if len(field_names) > 256:
                logger.warn('reducing the number of fields from %s to 256' %
                            len(field_names))
                field_names = field_names[:255]

            j = job.toJsonable()

            search = j['request']['search'] + ' | head %s ' % count
            params = {
                'required_field_list': '*',
                'status_buckets': str(j.get('statusBuckets', 0)),
                'remote_server_list': '*',
                'output_mode': 'json'
            }

            if j.get('reportSearch') and int(j.get('resultCount')) <= count:
                use_sid = sid
                etype = 'results'
            elif int(j.get('eventAvailableCount')) >= count:
                use_sid = sid
                etype = 'events'
            else:
                if j.get('searchEarliestTime'):
                    params['earliest_time'] = j['searchEarliestTime']
                if j.get('searchLatestTime'):
                    if not j['request'].get('latest_time'):
                        params['latest_time'] = j.get('createTime')
                    else:
                        params['latest_time'] = j['searchLatestTime']

            earliestTime = j.get('earliestTime')
            latestTime = j.get('latestTime')
            if latestTime is None:
                latestTime = j.get('createTime')
            try:
                response = fetcher.get(use_sid, etype, search, params)
                wb_client.add_data(
                    response, field_names, earliestTime, latestTime)
            except Exception, ex:
                logger.exception(ex)
                raise

        cherrypy.session.release_lock()

        BUF_SIZE = 1024 * 5

        def stream():
            ''' APP-815 - Use streaming mode and immediately start the download.
                Immediately send the first 8-bytes of the Excel XLS header since
                PEP-333 requires wsgi write() not send any headers until at least
                one byte is available for the response body.
            '''
            tmp_filename = "/tmp/excel_export_" + token + ".xlsx"
            wb_client.save(tmp_filename)
            f = open(tmp_filename, 'rb')
            data = f.read(BUF_SIZE)
            while len(data) > 0:
                yield data
                data = f.read(BUF_SIZE)
            os.remove(tmp_filename)

        return stream()

    def get_processed_job(self, sid, search=None):
        """ retrieve finished (optionally postprocessed) job """

        try:
            job = splunk.search.getJob(
                sid, sessionKey=cherrypy.session['sessionKey'])
        except:
            raise cherrypy.HTTPError('400', 'sid not found')

        if search:
            job.setFetchOption(search=search)

        while not job.isDone:
            time.sleep(.1)

        return job
