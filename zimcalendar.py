import time
import datetime
import calendar
import pyzim
import xml
from xml.dom.ext import PrettyPrint

class Appointment(object):
    def __init__(self, name = None, start = None, duration = None, xml = None):
        if xml:
            self._from_xml(xml)
        else:
            self.name = name
            self.start = start
            self.duration = duration

    def _get_context(self, doc):
        c = xml.xpath.CreateContext(doc)
        c.setNamespaces({"soap" : 'http://www.w3.org/2003/05/soap-envelope'})
        return c


    def _from_xml(self, doc):
        '''
        converts a xml node into an appointment
        '''
        self.start = None
        self.end = None
        self.doc = doc

        # First check what type of xml is sent. Zimbra can send different ones.

        c = self._get_context(doc)
        e = xml.xpath.Compile('//appt/inv/comp')
        results = e.evaluate(c)

        if len(results) > 0:
            # This is a full listing
            comp = results[0]

            self.name = comp.attributes['name'].value
            self.date = comp.attributes['d'].value
            self.id = doc.attributes['id'].value
            
            try:
                self.all_day = comp.attributes['allDay'].value
            except KeyError:
                self.all_day = 0
            if self.all_day == 0:
                e = xml.xpath.Compile('//appt/inv/comp/s')
                results = e.evaluate(c)
                a = results[0]
                self.start = a.attributes['d'].value

                e = xml.xpath.Compile('//appt/inv/comp/e')
                results = e.evaluate(c)
                e = results[0]
                self.end = a.attributes['d'].value

        else:
            # Probably a search result
            self.name = doc.attributes['name'].value
            self.date = doc.attributes['d'].value
            self.id = doc.attributes['id'].value
            
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

    def get_appointment(self, id):
        doc = self.build_soap_envelope()
        body = doc.getElementsByTagName('soap:Body')[0]
        search_request = doc.createElementNS('urn:zimbraMail', 'GetAppointmentRequest')
        search_request.setAttribute('id', str(id))
        body.appendChild(search_request)

        # Zimbra sends different xml when you only request a singular event


        doc = self._send_request(doc)

        try:
            appt = doc.getElementsByTagName('appt')[0]
        except IndexError:
            return False

        appt = Appointment(xml = appt)
        return appt

    def create_appointment(self):
        raise NotImplementedError

    def modify_appointment(self):
        raise NotImplementedError

    def cancel_appointment(self):
        raise NotImplementedError

    def get_free_or_busy(self):
        raise NotImplementedError

    def get_recurance(self):
        raise NotImplementedError

    def check_recurance_conficts(self):
        raise NotImplementedError

    def get_ical(self):
        raise NotImplementedError

    def send_invite_reply(self):
        raise NotImplementedError

    def import_appointment_request(self):
        raise NotImplementedError

    def dismiss_calendar_item_alarm(self):
        raise NotImplementedError

    def get_mini_cal_request(self):
        raise NotImplementedError

    
