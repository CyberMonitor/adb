from adversary.underscore_crew_201806 import generate_vba
from adversary.underscore_crew_201806 import payloads
from internals import office_doc_builder

playbook = generate_vba.generate(payloads.payload_list[0])

office_doc_builder.build_doc(template='C:\\Users\\h\\PycharmProjects\\adb\\adversary\\sample\\template.docx',
                             outfile='C:\\Users\\h\\Desktop\\outfile.doc',
                             actions=playbook)
