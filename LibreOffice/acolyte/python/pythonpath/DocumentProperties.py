import sys, uno

sys.stderr = sys.stdout
DOC_ID_NAME = "Acolyte Document Id"


class DocumentProperties:
    def __init__(self, doc):

        self.doc = doc
        self.doc_id = self.DocumentGuid
        self.properties = self.doc.DocumentProperties

    @property
    def DocumentGuid(self) -> str:
        udp = self.properties.UserDefinedProperties

        if not udp.getPropertySetInfo().hasPropertyByName(DOC_ID_NAME):
            udp.addProperty(DOC_ID_NAME ,128,"")
            udp.setPropertyValue(DOC_ID_NAME, str(uuid.uuid4()))

        return udp.getPropertyValue(DOC_ID_NAME)

    @property
    def Language(self) -> str:
        return self.doc.CharLocale

    @property
    def Title(self) -> str:
        return self.properties.Title

    @property
    def Subject(self) -> str:
        return self.properties.Subject

    @property
    def Characters(self) -> int:
        return self.doc.CharacterCount