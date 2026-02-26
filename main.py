from fastapi import FastAPI
from pydantic import BaseModel
from excel_service import create_project, add_update, get_updates, edit_update

app = FastAPI()

class ProjectRequest(BaseModel):
    project_name: str

class UpdateRequest(BaseModel):
    project_name: str
    date: str
    status: str

class FetchRequest(BaseModel):
    project_name: str
    date: str

class EditRequest(BaseModel):
    project_name: str
    date: str
    entry_number: int
    new_text: str

@app.post("/create-project")
def api_create_project(request: ProjectRequest):
    return {"message": create_project(request.project_name)}

@app.post("/add-update")
def api_add_update(request: UpdateRequest):
    return {"message": add_update(request.project_name, request.date, request.status)}

@app.post("/get-updates")
def api_get_updates(request: FetchRequest):
    return {"updates": get_updates(request.project_name, request.date)}

@app.post("/edit-update")
def api_edit_update(request: EditRequest):
    return {"message": edit_update(
        request.project_name,
        request.date,
        request.entry_number,
        request.new_text
    )}

@app.get("/")
def root():
    return {"message": "Status Management API is running"}