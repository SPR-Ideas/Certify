from itertools import count
from venv import create
import yaml
from certify import start , MAPPED_LOG , CERTIFCATE_DIR ,error , get_details
import os
from GoogleApi.gmail_api import  send_mail_for_participants
YAML = "sample_case/config.yaml"


def read_config():
    data : dict
    with open(YAML, "r") as stream:
        try:
            data= yaml.safe_load(stream)
        except yaml.YAMLError as exc:
            print(exc)
    if not data:
        error("Empty Config file")
    return data


def make_certificate(data):
    global CERTIFCATE_DIR
    # mapped_values :dict
    mapped_values = { "{{"+key+"}}" : value for key,value in data['variables'].items()}
    start(data['template'],
    data['data'],
    # data['variables'],
    mapped_values,
    data['save_location']
    )
    # CERTIFCATE_DIR  = data["save_location"]


def mail_wapper(raw_data,subject,content_file):

    if not os.path.exists(content_file):
        error(" Corrupted Content file for email")

    with open(content_file,'r') as fp:
        content = fp.read(4096)

    send_mail_for_participants(raw_data,subject,content)


def make_dynamic_content(raw_data,content,mapped_values,csv_file):
    """Making the dynamic content and append it to raw_data."""
    details = get_details(csv_file,list(mapped_values.values()))
    count = 1
    for csv_iterator in details.iterrows():

        find_replace = {

                key:csv_iterator[1][values] for key,values in mapped_values.items()
            }
        con = content
        for search_str , rep_str in find_replace.items():
            con = con.replace(search_str,rep_str)

        raw_data[count] = raw_data[count].append(con)
        count+=1

    return raw_data



def compile_config(file=YAML):
    """
        It complies the config file and return the data.
    """
    if os.path.exists(file):
        try :
            data = read_config()
            if "Certify" not in data:
                error("""Error in YAML file
                The Header 'Certify:' is not mentioned.
                """ )
            else:
                data = data["Certify"]
            err_no =1
            er = "Compliation terminated \n\n"

            required_feild = [
                'template',"data",'save_location' ,"variables",
                "sent_email" ,"content" ,"subject"
            ]

            feilds_not_there = [i for i in required_feild if i not in data]

            # if sent_email is not there then it set false.
            # if sent_email is false then content and subject feild are not-nesscceary
            if "sent_email" in feilds_not_there:
                feilds_not_there.remove("sent_email")
                if "content" in feilds_not_there:
                    feilds_not_there.remove("content")
                if "subject" in  feilds_not_there:
                    feilds_not_there.remove("subject")
                data['sent_email'] = False


            if feilds_not_there:

                for i in feilds_not_there:
                    er+="{}. {} feild not there \n".format(err_no,i)
                    err_no+=1
                error(er)
            return data
        except Exception as e :
            error("Courpted yaml file {}".format(e))

    else:
        error("No config.ymal file >> create one. ")



def assign_task():
    """
        This function gonna assign the task based on the config file.
    """

    data = compile_config()
    if data:
        # mail_data = data['Certify']
        make_certificate(data)
        if "sent_email" in data:

            if data["sent_email"]:
                raw_data = {}
                counter =1

                for key , value in MAPPED_LOG.items():
                    raw_data[counter] = [key,value]
                    counter+=1

                mail_wapper(raw_data , data['subject'],data['content'])



def run_sand_box():
    global YAML
    YAML = "sample_case/config.yaml"
    data = compile_config()
    # print(data)
    make_certificate(data)


if __name__ == "__main__":
    assign_task()
    # run_sand_box()