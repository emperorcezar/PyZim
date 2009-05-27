from xml.dom import implementation
from xml.dom.ext import PrettyPrint
from xml.dom.ext import Print
from httplib import HTTPSConnection
import StringIO
from xml.dom import minidom
from xml import xpath
import xml

class PyZim(object):

    def __init__(self, server):
        if server == '' or not server:
            raise ValueError, "Server is undefined"
        self.server = server
        self.auth_token = ''
        self.session_id = ''
        self._last_response = None
        self.username = None

    def build_soap_envelope(self):
        namespace = 'http://www.w3.org/2003/05/soap-envelope'

        # create XML DOM document
        doc = implementation.createDocument(None, '', None)

        # create soap envelope element with namespaces 
        soapenv = doc.createElementNS(namespace, "soap:Envelope")

        # add soap envelope element
        doc.appendChild(soapenv)

        # create header element
        header = doc.createElementNS(namespace, "soap:Header")
        soapenv.appendChild(header)

        context = doc.createElementNS("urn:zimbra", "context")
        header.appendChild(context)

        # If we have and authToken, use it
        if self.auth_token != '':
            auth_token = doc.createElement('authToken')
            auth_token.appendChild(doc.createTextNode(self.auth_token))
            context.appendChild(auth_token)

        if self.session_id != '':
            session_id = doc.createElement('sessionId')
            session_id.setAttribute('id', self.session_id)

        body = doc.createElement('soap:Body')
        soapenv.appendChild(body)

        #PrettyPrint(doc)
        return doc

    def _send_request(self, doc):
        con = HTTPSConnection(self.server)

        soapString = StringIO.StringIO()
        PrettyPrint(doc, soapString)
        soapString = soapString.getvalue()

        con.request('POST', '/service/soap', soapString)
        res = con.getresponse()
        response = res.read()
        doc = minidom.parseString(response)

        # Check for session id and change id and set them
        
        c = self._get_context(doc)
        e = xml.xpath.Compile('//context/sessionId/text()')
        result = e.evaluate(c)
        if len(result) > 0:
            self.session_id = (result[0]).data


        e = xml.xpath.Compile('//context/change/@token')
        result = e.evaluate(c)
        if len(result) > 0:
            self.change_id = (result[0]).value

        self._last_response = doc

        return doc
    
    def authenticate(self, username, password):
        doc = self.build_soap_envelope()

        body = doc.getElementsByTagName('soap:Body')[0]

        auth_request = doc.createElementNS('urn:zimbraAccount', 'AuthRequest')
        account = doc.createElement('account')
        account.setAttribute('by', 'name')
        account.appendChild(doc.createTextNode(username))
        password_node = doc.createElement('password')
        password_node.appendChild(doc.createTextNode(password))

        auth_request.appendChild(account)
        auth_request.appendChild(password_node)
        body.appendChild(auth_request)

        doc = self._send_request(doc)
        
        c = self._get_context(doc)
        e = xml.xpath.Compile('//soap:Body/AuthResponse/authToken/text()')

        try:
            self.auth_token = (e.evaluate(c)[0]).data
        except IndexError:
            e = xml.xpath.Compile('//soap:Body/soap:Fault/soap:Reason/soap:Text/text()')
            self.error_text = (e.evaluate(c)[0]).data
            return False
        
        self.username = username

        return True


    def _get_context(self, doc):
        c = xml.xpath.CreateContext(doc)
        c.setNamespaces({"soap" : 'http://www.w3.org/2003/05/soap-envelope'})
        return c
            
