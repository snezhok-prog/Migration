from copy import deepcopy
from _config import APPEAL_SETTINGS

SUBJECT_UL = {
    "guid": None,
    "kind": {
        "type": "participant",
        "name": "Участник",
        "subKinds": [
            {
                "name": "Юридическое лицо",
                "specialTypeId": "ulApplicant",
                "headerOptions": [
                    "f|data.organization.shortName",
                    "s|, ОГРН:",
                    "f|data.organization.ogrn"
                ],
                "shortHeaderOptions": [
                    "f|data.organization.shortName"
                ],
                "shortCode": "ul"
            },
            {
                "name": "Индивидуальный предприниматель",
                "specialTypeId": "ipApplicant",
                "headerOptions": [
                    "f|data.person.lastName",
                    "s| ",
                    "f|data.person.firstName",
                    "s| ",
                    "f|data.person.middleName",
                    "s|, ",
                    "s|ОГРН:",
                    "f|data.person.ogrn"
                ],
                "shortHeaderOptions": [
                    "f|data.person.lastName",
                    "s| ",
                    "f|data.person.firstName|1",
                    "s|.",
                    "f|data.person.middleName|1",
                    "s|."
                ],
                "shortCode": "ip"
            }
        ],
        "subKind": {
            "name": "Юридическое лицо",
            "specialTypeId": "ulApplicant",
            "headerOptions": [
                "f|data.organization.shortName",
                "s|, ОГРН:",
                "f|data.organization.ogrn"
            ],
            "shortHeaderOptions": [
                "f|data.organization.shortName"
            ],
            "shortCode": "ul"
        }
    },
    "data": {
        "person": {},
        "organization": {
            "opf": {
                "_id": "6836ecb3132ead24cba36372",
                "git": "1082",
                "auid": 0,
                "code": "10000",
                "guid": "0fe96ac2-25d2-138e-fb01-47264567d80c",
                "name": "ОРГАНИЗАЦИОННО-ПРАВОВЫЕ ФОРМЫ ЮРИДИЧЕСКИХ ЛИЦ, ЯВЛЯЮЩИХСЯ КОММЕРЧЕСКИМИ КОРПОРАТИВНЫМИ ОРГАНИЗАЦИЯМИ",
                "dateCreation": None,
                "userCreation": {},
                "parentEntries": "catalogueOpf",
                "dateLastModification": None,
                "userLastModification": {}
            },
            "shortName": "Агентство \"Полилог\"",
            "name": "Агентство \"Полилог\"",
            "ogrn": "0000000000000",
            "inn": "0000000000",
            "kpp": "000000000",
            "registrationAddress": {
                "fullAddress": "Приморский край, г. Владивосток, ул. Береговая, д. 8, кв. 2"
            }
        },
        "personInOrg": {
            "position": None,
            "authority": None
        }
    },
    "specialTypeId": "ulApplicant",
    "parentEntries": f"{APPEAL_SETTINGS['parentEntries']}.subjects",
    "xsdData": {
        "phone": None,
        "email": None,
        "factAddress": {
            "fullAddress": None
        }
    },
    "mainXsdDataValid": True,
    "xsdDataValid": True,
    "subjectTypeXsdDataValid": True,
    "objectTypeXsdDataValid": True,
    "shortHeader": None,
    "header": None,
    "mainId": None,
    "entityType": "subjects",
    "name": None
}

SUBJECT_IP = {
    "guid": None,
    "kind": {
        "type": "participant",
        "name": "Участник",
        "subKinds": [
            {
                "name": "Юридическое лицо",
                "specialTypeId": "ulApplicant",
                "headerOptions": [
                    "f|data.organization.shortName",
                    "s|, ОГРН:",
                    "f|data.organization.ogrn"
                ],
                "shortHeaderOptions": [
                    "f|data.organization.shortName"
                ],
                "shortCode": "ul"
            },
            {
                "name": "Индивидуальный предприниматель",
                "specialTypeId": "ipApplicant",
                "headerOptions": [
                    "f|data.person.lastName",
                    "s| ",
                    "f|data.person.firstName",
                    "s| ",
                    "f|data.person.middleName",
                    "s|, ",
                    "s|ОГРН:",
                    "f|data.person.ogrn"
                ],
                "shortHeaderOptions": [
                    "f|data.person.lastName",
                    "s| ",
                    "f|data.person.firstName|1",
                    "s|.",
                    "f|data.person.middleName|1",
                    "s|."
                ],
                "shortCode": "ip"
            }
        ],
        "subKind": {
            "name": "Индивидуальный предприниматель",
            "specialTypeId": "ipApplicant",
            "headerOptions": [
                "f|data.person.lastName",
                "s| ",
                "f|data.person.firstName",
                "s| ",
                "f|data.person.middleName",
                "s|, ",
                "s|ОГРН:",
                "f|data.person.ogrn"
            ],
            "shortHeaderOptions": [
                "f|data.person.lastName",
                "s| ",
                "f|data.person.firstName|1",
                "s|.",
                "f|data.person.middleName|1",
                "s|."
            ],
            "shortCode": "ip"
        }
    },
    "data": {
        "person": {
            "citizenship": {
                "code": "RUS",
                "name": "Россия"
            },
            "matchesRegistrationAddress": {
                "tempRegistrationAddress": False,
                "factAddress": False
            },
            "lastName": None,
            "firstName": None,
            "middleName": None,
            "declensionOfFIO": {
                "fullString": None
            },
            "birthday": {
                "date": {
                    "year": 1984,
                    "month": 12,
                    "day": 31
                },
                "jsDate": "1984-12-30T21:00:00.000Z",
                "formatted": "31.12.1984",
                "epoc": 473288400
            },
            "documentType": [
                {
                    "id": "59",
                    "text": "Паспорт гражданина РФ, удостоверяющего личность за пределами РФ"
                }
            ],
            "documentSeries": "12",
            "documentNumber": "1231231",
            "documentIssueDate": "2025-12-12T00:00:00.000+0300",
            "documentIssuer": {
                "_id": "59d3925d1ccf6d8a5605b77b",
                "code": "023-009",
                "guid": "d45d77b5-6bb3-4320-bfdf-3321d848ea2b",
                "name": "РЕЗЕРВ РЕСПУБЛИКИ БАШКОРТОСТАН",
                "dateEnd": "null",
                "idValue": 3114
            },
            "registrationAddress": {
                "fullAddress": "Приморский край, г. Владивосток, ул. Береговая, д. 8, кв. 2"
            },
            "factAddress": {
                "fullAddress": "Приморский край, г. Владивосток, ул. Береговая, д. 8"
            },
            "inn": "000000000000",
            "mobile": "+7 (111) 111 11 11",
            "email": "pochta@pochta.pochta",
            "ogrn": "322440000000311"
        }
    },
    "specialTypeId": "ipApplicant",
    "parentEntries": f"{APPEAL_SETTINGS['parentEntries']}.subjects",
    "agreeMkguInterview": True,
    "xsdData": {
        "nameIP": None,
        "birthPlace": None
    },
    "mainXsdDataValid": True,
    "xsdDataValid": True,
    "subjectTypeXsdDataValid": True,
    "objectTypeXsdDataValid": True,
    "shortHeader": None,
    "header": None
}
