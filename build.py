from internals import office_doc_builder
from internals import random_stuff
from internals import utils
import random
import importlib
import os

def build_files(adversary, out_dir, count=1, filetype="doc", extension="doc", debug=False):

    generate_vba = importlib.import_module("adversary." + adversary + ".generate_vba")
    payloads = importlib.import_module("adversary." + adversary + ".payloads")

    #  set these

    out_root_path = out_dir

    c = 0

    filetype_and_extension = utils.reconcile_extension_and_format(extension=extension, filetype=filetype)

    while c < count:
        playbook = generate_vba.generate(random.choice(payloads.payload_list))
        if 'change_extension_after_save' in filetype_and_extension:
            playbook.append({'change_extension_after_save': filetype_and_extension['change_extension_after_save']})
        playbook.append({'set_save_format': filetype_and_extension['set_save_format']})
        out_file_name = random_stuff.gen_doc_name(filetype_and_extension['set_save_extension'])
        out_file_path = os.path.join(out_root_path, out_file_name)
        initials, name = random_stuff.getrandomauthor()
        utils.set_doc_author_keys(userinitials=initials, username=name)
        playbook.append({'set_author': name})
        print("[*] Building document " + out_file_name + " with author: " + name)
        if debug == True:
            print("* Playbook: ")
            print(playbook)
            print("")
        office_doc_builder.build_doc(template=os.path.join(os.path.join(os.path.split(__file__)[0], 'adversary', adversary, 'template.docx')),
                                     outfile=out_file_path,
                                     actions=playbook)
        c = c + 1


if __name__ == '__main__':
    out_directory = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    #build_files('underscore_crew_201806', out_dir=out_directory, count=5)
    build_files('sample_with_network_test', out_dir=out_directory, count=1, filetype="docm", extension="rtf")
