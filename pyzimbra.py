from xml.dom import implementation
from xml.dom.ext import PrettyPrint
from xml.dom.ext import Print
from httplib import HTTPSConnection
import StringIO
from xml.dom import minidom
from xml import xpath
import xml

class PyZimbra(object):

    def __init__(self, server):
        if server == '' or not server:
            raise ValueError, "Server is undefined"
        self.server = server

    def build_soap_header(self):
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

        body = doc.createElement('soap:Body')
        soapenv.appendChild(body)

        #PrettyPrint(doc)
        return doc

    def authenticate(self, username, password):
        doc = self.build_soap_header()

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

        soapString = StringIO.StringIO()
        
        PrettyPrint(doc, soapString)
        
        soapString = soapString.getvalue()

        con = HTTPSConnection(self.server)

        con.request('POST', '/service/soap', soapString)
        res = con.getresponse()
        response = res.read()
        doc = minidom.parseString(response)
        #PrettyPrint(doc_node)

        #return doc_node

        c = self._get_context(doc)
        e = xml.xpath.Compile('//soap:Body/AuthResponse/authToken/text()')
        self.auth_token = (e.evaluate(c)[0]).data
        e = xml.xpath.Compile('//soap:Body/AuthResponse/sessionId/text()')
        self.auth_token = (e.evaluate(c)[0]).data


    def _get_context(self, doc):
        c = xml.xpath.CreateContext(doc)
        c.setNamespaces({"soap" : 'http://www.w3.org/2003/05/soap-envelope'})
        return c
            