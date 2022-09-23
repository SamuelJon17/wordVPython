import os
import streamlit as st
from docx import Document
import re

# Reference
# https://docs.streamlit.io/library/api-reference/session-state
# https://discuss.streamlit.io/t/how-to-get-multiple-inputs-using-same-text-input-box/28435/2 

st.title('Find And Replace Batch Process')
st.subheader('Upload Word Document(s) for Processing')
files = st.file_uploader('Multiple Allowed', accept_multiple_files=True)
path = st.text_input("Please Enter Path to Save", placeholder = "/User/Desktop/Downloads/", key = "path")

if "findReplace" not in st.session_state:
    st.session_state.findReplace = []

def findReplaceExact(String, find, replace, caseSens = False):
    """
    Find and replace text in String
    @param String: String from either a Paragraph or Runner object
    @param find: Substring wanting to replace
    @param replace: Replacement string
    @param caseSens: True - if case sensitive, False otherwise
    @return String, original string if find not in String otherwise updated String with replace inserted
    """
    if not caseSens:
        regex = r"(?i)(\b" + str(find) + r"\b)"
        if len(re.findall(regex, String))>0:
            result = re.split(regex, String)
            return ''.join([replace if i.lower() == find.lower() else i for i in result])
        else:
            return String
    else:
        regex = r"(\b" + str(find) + r"\b)"
        if len(re.findall(regex,String))>0:
            result = re.split(regex, String)
            return ''.join([replace if i.lower() == find.lower() else i for i in result])
        else:
            return String

def findReplaceExactOverDocument(document, find, replace, caseSens = False):
    """
    Find and replace text in Entire document
    @param document: Document object
    @param find: Substring wanting to replace
    @param replace: Replacement string
    @param caseSens: True - if case sensitive, False otherwise
    @return String, original string if find not in String otherwise updated String with replace inserted
    """
    for paragraph in document.paragraphs:
        if findReplaceExact(paragraph.text, find, replace, caseSens=caseSens) != paragraph.text:
            st.write("\n=============================")
            st.write("**The following statement:** ")
            st.write(f"\n\t'{paragraph.text}'")
            for run in paragraph.runs:
                if (findReplaceExact(run.text, find, replace, caseSens=caseSens) != run.text):
                    run.text = findReplaceExact(run.text, find, replace, caseSens=caseSens)
            st.write(f"**Was replaced to:**")
            st.write(f"\n\t'{paragraph.text}'")
    pass

def add_callback():
    # Handle if find already in list
    findIndex = [i for i, x in enumerate(st.session_state.findReplace) if x['find'] == st.session_state.find]
    if len(findIndex) >0:
        st.session_state.findReplace[findIndex[0]].update({"find":st.session_state.find, "replace":st.session_state.replace,
        "caseSensitive":st.session_state.case})
    
    # Otherwise update list
    else:
        st.session_state.findReplace.append({"find":st.session_state.find, "replace":st.session_state.replace,
        "caseSensitive":st.session_state.case})
    

def delete_callback():
    # Remove last added index
    if len(st.session_state.findReplace)==0:
        pass
    else:
        st.session_state.findReplace = st.session_state.findReplace[:-1]

def saveDocs(documents, docName):
    if len(documents) ==0:
        st.subheader("Please upload a document")
    else:
        for i in range(len(documents)):
            docName = docName[i].split(".docx")[0] + "_Processed.docx"
            documents[i].save(os.path.join(path,docName))
            st.write(f"**File successfully saved at: {os.path.join(path,docName)}**")
        st.subheader("Please refresh page to start a new session".title())
    pass

## Update state session find/replace statements
with st.form(key='myForm'):
    st.write("Find and Replace Documents")
    find = st.text_input("Find", placeholder="Add find word", key = "find")
    replace = st.text_input("Replace", placeholder = "Add replace word", key = "replace")
    caseSensitive = st.checkbox("Case Sensitive", value = False, key = "case")
    submit_button = st.form_submit_button(label='Add', on_click=add_callback)
    delete_button = st.form_submit_button(label="Remove Last", on_click=delete_callback)

## Print state session find/replace statements
count=1
for obj in st.session_state.findReplace:
    st.write(f"{count}. Find: {obj['find']}  | Replace: {obj['replace']}  | Case Sesnsitive: {obj['caseSensitive']}\n")
    count+=1


## Now: This will loop through each document and within each document loop again for each findReplace item
## Future: This can be improved by looping through document only one pass through and check all findReplace at once
## Future:  Save JSON of "Find/Replace/CaseSens" to import rather than typing it in each time
if st.button("Process Documents"):
        if files is not None:
            docNames=[]
            documents = []
            for document in files:
                docName = str(document.name)
                docNames.append(docName)
                st.subheader(f"{docName}")
                document = Document(document)
                documents.append(document)
                for item in st.session_state.findReplace:
                    findReplaceExactOverDocument(document, item['find'], item['replace'], item['caseSensitive'])
        saveDocs(documents, docNames)
        