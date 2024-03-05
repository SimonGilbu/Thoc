# Summary
The goal of Thoc is to provide a set of useful tools for developers that increases 
speed and productivity when working with PowerPoint. 

# Techinal implementation
Thoc like most taskpane PPT Web-adins written with the Office JS API consists of 
two primary parts, the XML manifest and the "application", the XML manifest is a 
detailed description of how the addin interaces with PPT, the application is the 
collective term for the logic used in creating a web-app that is run inside the 
context of PPT

### The application
The Thoc application is primarily utilizing a pure JS/TS frontend, 
any server could be used to serve the relevant content. from a static nginx dir 
response to templateing with Tera or Jinja using rust or python respectively. 
for the case of this application all development has been done using a simple 
Flask Python server. 

### Technology choices
to ease development the application uses the typescript definitions for the office api, 
and compiles to static JS using TSC, it also does styling with Tailwindcss,

i think the npm pacakge for the types is something like `@types/office-js`.

```bash
# to run TSC
$ npx tsc .\static\js\thoc.ts [--watch]

# to run tailwind
$ npx tailwindcss -i .\static\css\main.css -o .\static\css\tailwind.css [--watch]

# to run the server on windows [assuming venv]
.\Scripts\activate
python server.py
```
### Requirements
you need a version of both Node and Python,
a shared folder for testing the manifest file,
and ofcrouse a modern version of powerpoint installed
locally
