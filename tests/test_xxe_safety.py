"""XXE-гарантия парсера python-docx зафиксирована тестом (аудит 2.0.1, P3).

Безопасность держится на lxml-парсере python-docx с resolve_entities=False.
Раньше это знание жило только в комментарии; теперь тест падает, если
настройка парсера изменилась (обновление python-docx и т. п.).
"""
from __future__ import annotations

import inspect

XXE_XML = (
    '<?xml version="1.0"?>'
    "<!DOCTYPE w:document [<!ENTITY secret 'XXE_LEAKED_VALUE'>]>"
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body><w:p><w:r><w:t>&secret;</w:t></w:r></w:p></w:body></w:document>"
)


def test_docx_parser_disables_entity_resolution():
    from docx.oxml import parser as docx_parser

    source = inspect.getsource(docx_parser)
    assert "resolve_entities=False" in source, (
        "python-docx собран/обновлён с resolve_entities=True — XXE-защита "
        "импортёра нарушена; обновление пакета требует ревью"
    )


def test_entities_are_not_expanded_in_document_parts():
    from docx.oxml.parser import parse_xml

    element = parse_xml(XXE_XML)
    serialized = element.xml if hasattr(element, "xml") else str(element)
    assert "XXE_LEAKED_VALUE" not in serialized, (
        "сущность DOCTYPE разворачивается при разборе .docx — XXE-уязвимость"
    )
