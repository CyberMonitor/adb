from internals import office_doc_builder
from internals import random_stuff
from internals import utils
import importlib
import os

def build_files(adversary, out_dir, count=1):

    generate_vba = importlib.import_module("adversary." + adversary + ".generate_vba")
    payloads = importlib.import_module("adversary." + adversary + ".payloads")

    #  set these

    out_root_path = out_dir

    playbook = generate_vba.generate(payloads.payload_list[0])

    c = 0

    while c < count:
        out_file_name = random_stuff.gen_doc_name("doc")
        out_file_path = os.path.join(out_root_path, out_file_name)
        initials, name = random_stuff.getrandomauthor()
        utils.set_doc_author_keys(userinitials=initials, username=name)
        playbook.append({'set_author': name})
        print("[*] Building document " + out_file_name + " with author: " + name)
        office_doc_builder.build_doc(template=os.path.join(os.path.join(os.path.split(__file__)[0], 'adversary', adversary, 'template.docx')),
                                     outfile=out_file_path,
                                     actions=playbook)
        c = c + 1


if __name__ == '__main__':
    out_directory = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    #build_files('underscore_crew_201806', out_dir=out_directory, count=5)
    build_files('sample_with_network_test', out_dir=out_directory, count=1)
