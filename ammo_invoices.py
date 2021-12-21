"""
AUTHOR: James Dixon - Durham university Rifle CLub - 2021
USE: for use in durham university rifle club

this is terrible hackish code created entirely for the DURC google form response CSVs.
will have to be edited to work with any other layout of csv
"""
from datetime import datetime
import xlrd
import pandas as pd
import hashlib
import numpy as np
import csv
import time
import math

names = []
membership_file = "data/DURC_membership_form.xlsx"
post_shoot_file = "data/Post_shoot_responses.xlsx"

ammo_prices = {'cci': 7, 'sport': 9, 'club': 12, 'match': 20, 'tenex': 23}
start_date = datetime.strptime("01/01/21", '%d/%m/%y')
end_date = datetime.strptime("01/01/22", '%d/%m/%y')
members = []
entrys = []


def readFileLine(path):
    """
    generates lines of the CSV line by line
    first line is generaly column names so should be discarded
    :param path: path to the CSV file
    :return: array of the line in the CSV
    """

    f = pd.read_excel(f"{path}", sheet_name=None)
    f = f[list(f.keys())[0]]
    for i, r in f.iterrows():
        yield r

def blancInvoice():
    return pd.DataFrame(columns=["timestamp", "Rounds Shot", "Ammo", "shoot cost"])


class Entry:
    def __init__(self, data):
        self.data = data

        self.time_stamp = self.data['Timestamp'].asm8
        self.name = self.data['Name']
        self.number_shot = self.data['number Rounds Shot']
        if not (type(self.number_shot) is float or type(self.number_shot) is int):
            self.number_shot = 15
        self.ammo_type = self.data['Ammo Type']

        self.dictKeys = self.keys()


        self.cost = self.price()

    def price(self):
        if not pd.isna(self.ammo_type) and not pd.isna(self.number_shot):
            p = ammo_prices[self.ammo_type.lower()]
            return p * self.number_shot /100
        return 0

    def formatDate(self, unformatted):
        parts = unformatted.split("/")
        full = ""
        for part in parts:
            if len(part) == 1:
                part = "0" + part + "/"
            else:
                part = part + "/"
            full = full + part
        full = full[:-1]
        try:
            dt = datetime.strptime(f"{full}", '%d/%m/%Y')
        except:
            dt = datetime.strptime(f"{full}", '%d/%m/%y')

    def keys(self):
        if not pd.isna(self.name):
            self.name = self.name.lower()
            self.name = self.name.split(' ')
            n = []
            for np in self.name:
                if np != "":
                    n.append(np)
            self.name = n

            return self.name
        return []


class Member:
    def __init__(self, data):
        self.data = data

        self.name = self.data['name']
        self.Timestamp = self.data['Timestamp']
        self.DateofBirth = self.data['Date of Birth']
        self.College = self.data['College']
        self.yearofstudy = self.data['year of study']
        self.Termaddress = self.data['Term address (FULL ADDRESS + POSTCODE) - Note: livers in need to include BUILDING and ROOM NUMBER']
        self.HomeAddress = self.data['Home address (FULL ADDRESS + POSTCODE)']

        if not pd.isna(self.data['Surname']):
            self.Surname = self.data['Surname'].lower().split(" ")
        else:
            self.Surname = self.data['Surname']

        if not pd.isna(self.data['Forename(s)']):
            self.Forename = self.data['Forename(s)'].lower().split(" ")
        else:
            self.Forename = self.data['Forename(s)']

        if not pd.isna(self.data['Durham Email [ending in @durham.ac.uk]']):
            self.DurhamEmail = self.data['Durham Email [ending in @durham.ac.uk]'].lower().split("@")
        else:
            self.DurhamEmail = self.data['Durham Email [ending in @durham.ac.uk]']

        self.entrys = []
        self.lookup = self.dictEntry()

        self.invoice = blancInvoice()
        self.total = 0
        self.referance = ""

    def ref(self):
        return str(hash((self.Surname[0] + str(self.total))))[-8:]

    def fillInvoice(self):
        for i in range(len(self.entrys)):
            shoot = self.entrys[i]
            self.invoice.loc[i] = [str(shoot.data["Timestamp"]), shoot.number_shot, shoot.ammo_type, shoot.cost]
        self.total = sum(self.invoice["shoot cost"])
        if self.total!=0:
            self.referance = self.ref()
            self.invoice.loc[len(self.entrys)] = [f"total to pay: £{self.total}" ] + [f"Bank transfer Reference: {self.ref()}"] + ["_" for _ in range(2)]

    def genInvoice(self):
        if self.total != 0:
            self.invoice.to_excel(f"invoices/{self.DurhamEmail[0]}_12-2021.xlsx",
                         sheet_name='term 1 2021')

    def memberOwes(self):
        fullName = ""
        for n in self.Forename + self.Surname:
            fullName += n
            fullName += " "
        return [fullName, self.DurhamEmail[0], self.total, self.referance]

    ['Name', 'Email', 'total', 'reference']
    def dictEntry(self):
        d = dict()

        kd = ""

        try:
            a = len(self.Forename)
            valid = True
        except:
            valid = False

        if valid:
            self.Forename = self.Forename



            kd = self.Surname[0]

            if kd != "":
                d[kd] = self

        return d


def createMembers(file):
    memberGen = readFileLine(f"{file}")
    for m in memberGen:
        members.append(Member(m))


def createEntrys(file):
    ErntryGen = readFileLine(f"{file}")
    for m in ErntryGen:
        entrys.append(Entry(m))


def assighnShoots():
    unassighned = []

    memberDict = dict()
    for m in members:
        memberDict.update(m.lookup)

    for shoot in entrys:
        possible_shooters = []
        found = False
        for i in range(len(shoot.dictKeys)):
            pk = shoot.dictKeys[len(shoot.dictKeys)-i-1]
            try:
                if not found:
                    possible_shooters.append(memberDict[pk])
                    found = True
            except:
                pass
        if len(list(set(possible_shooters))) != 0:
            list(set(possible_shooters))[0].entrys.append(shoot)

        else:
            unassighned.append(shoot)
    pass

def createInvoices():
    for m in members:
        m.fillInvoice()
        m.genInvoice()

    checkForm = pd.DataFrame(columns=['Name', 'Email', 'total', 'reference'])
    i = 0
    for m in members:
        if m.total != 0:
            checkForm.loc[i] = m.memberOwes()
            i+=1

    checkForm.to_excel(f"checklist_12-2021.xlsx",sheet_name='term 1 2021')



def main():
    createMembers(membership_file)
    createEntrys(post_shoot_file)
    assighnShoots()
    createInvoices()
    pass


if __name__ == '__main__':
    main()
