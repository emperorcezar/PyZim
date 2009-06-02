import time
import datetime
import calendar
import pyzim
from xml.dom.ext import PrettyPrint

class Appointment(object):
    def __init__(self, name = None, start = None, duration = None, xml = None):
        if xml:
            self._from_xml(xml)
        else:
            self.name = name
            self.start = start
            self.duration = duration

    def _from_xml(self, doc):
        '''
        converts a xml node into an appointment
        '''

        self.name = doc.attributes['name'].value
        self.date = doc.attributes['d'].value
        try:
            self.all_day = doc.attributes['allDay'].value
        except KeyError:
            self.all_day = 0
        if self.all_day == 0:
            self.duration = doc.attributes['dur'].value
            inst = doc.getElementsByTagName('inst')[0]
            self.start = inst.attributes['s'].value

        
        

class ZimCalendar(pyzim.PyZim):
    def init(self, server):
        super(ZimCalendar, self).__init__(server)
    
    def get_current_month(self):
        '''
        Get all appointments for the current month
        '''
        doc = self.build_soap_envelope()
        body = doc.getElementsByTagName('soap:Body')[0]

        search_request = doc.createElementNS('urn:zimbraMail', 'SearchRequest')

        now = datetime.datetime.now()

        start_date = time.mktime(time.strptime(str(now.month) + '/1/' + str(now.year) ,"%m/%d/%Y"))
        end_date = time.mktime(time.strptime(
            str(now.month) + '/' + str(calendar.monthrange(now.year, now.month)[1]) + '/' + str(now.year) ,"%m/%d/%Y")
                                     )

        start_date = int(start_date * 1000)
        end_date = int(end_date * 1000)

        search_request.setAttribute('calExpandInstStart', str(start_date))
        search_request.setAttribute('calExpandInstEnd', str(end_date))

        search_request.setAttribute('types', 'appointment,message,task')

        query = search_request.appendChild(doc.createElement('query'))
        query.appendChild(doc.createTextNode('inid:' + str(self._calendar_id)))

        body.appendChild(search_request)

        doc = self._send_request(doc)

        appts = doc.getElementsByTagName('appt')

        if len(appts) == 0:
            return []

        appointments = []
        for appt in appts:
            appointments.append(Appointment(xml = appt))

        return appointments

    
