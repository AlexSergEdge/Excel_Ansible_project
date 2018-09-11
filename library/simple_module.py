from ansible.module_utils.basic import *

def run_module():
    fields = dict(
        name = dict(required=True, type='str'),
        secret = dict(required=True, type='str'),
    )
    
    result = dict(
        changed=False,
        message=''
    )
    
    module = AnsibleModule(argument_spec=fields, supports_check_mode=True)
    
    if module.check_mode:
        return result
    
    # get all input data
    name = module.params['name']
    secret = module.params['secret']
    result['failed'] = False
    result['changed'] = False
    result['message'] = "Successfully called simple module"
    module.exit_json(**result)


def main():
    run_module()

if __name__ == '__main__':
    main()