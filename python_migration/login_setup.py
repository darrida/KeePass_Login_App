#!/usr/bin/env python
# coding: utf-8

# Python Keepass Login Setup

# In[2]:


import getpass
from passlib.hash import pbkdf2_sha256
try:
    from configparser import ConfigParser
except ImportError:
    from ConfigParser import ConfigParser  # ver. < 3.0


# In[7]:



def createini(password, username):
    ini_file_name = 'kp_config.ini' 
    hash_name = username + "_hash"
    
    # instantiate
    config = ConfigParser()
    
    # parse existing file
    with open(ini_file_name, 'w') as configfile:
        config.write(configfile)

    # parse existing file
    config.read(ini_file_name)

    # add a new section and some values
    config.add_section('ENTRY 1')
    config.set('ENTRY 1', hash_name, password)
    #config.set('ENTRY 1', 'user1', username)

    with open(ini_file_name, 'w') as configfile:
        config.write(configfile)
        
def verify(username):
    pw_verify = getpass.getpass("Confirm your password: ")
    hash_name = username + "_hash"
    
    # instantiate
    config = ConfigParser()
    
    # parse existing file
    config.read('kp_config.ini')
    
    # read values from a section
    read_hash = config.get('ENTRY 1', hash_name)
    #read_user = config.get('ENTRY 1', 'user1')
    #int_val = config.getint('section_a', 'int_val')
    #float_val = config.getfloat('section_a', 'pi_val')
    i = 0
    while ((pbkdf2_sha256.verify(pw_verify, read_hash) == False) & (i < 5)):
        print("\nUsername or password incorrect, please try again.")
        pw_verify = getpass.getpass("Confirm your password: ")
        i += 1
        #while (pbkdf2_sha256.verify(pw_verify, read_hash) == False):
    if (pbkdf2_sha256.verify(pw_verify, read_hash) == True):       
        print("\nThanks " + username + ", your password was successful.")
    else:
        print("Too many tries. Please start over.")
        
def login(username, password):
    hash_name = username + "_hash"
    
    # instantiate
    config = ConfigParser()
    
    # parse existing file
    config.read('kp_config.ini')
    
    # read values from a section
    read_hash = config.get('ENTRY 1', hash_name)

    i = 0
    while ((pbkdf2_sha256.verify(password, read_hash) == False) & (i < 5)):
        print("\nUsername or password incorrect, please try again.")
        password = getpass.getpass("Confirm your password: ")
        i += 1
        #while (pbkdf2_sha256.verify(pw_verify, read_hash) == False):
    if (pbkdf2_sha256.verify(password, read_hash) == True):       
        print("\n\nThanks " + username + " -- login successful.")
    else:
        print("Too many tries. Please start over.")
        
def main():
    user = input("What is your username?")
    password = getpass.getpass("What is your password?")
    #password = input("What is your password?")
    hash = pbkdf2_sha256.hash(password, rounds=500000, salt_size=16)
    createini(hash, user)
    #pw_verify = input("Confirm your password")
    verify(user)


# In[8]:


main()


# In[10]:


login_user = input("Username: ")
login_pw = getpass.getpass("Password: ")
login(login_user, login_pw)