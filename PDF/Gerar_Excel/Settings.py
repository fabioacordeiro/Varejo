{
  "window.zoomLevel": 2,
  "editor.fontSize": 16,
  "editor.hover.enabled": true,
  "workbench.startupEditor": "none",
  "explorer.compactFolders": false,
  "terminal.integrated.fontSize": 16,
  "editor.rulers": [80, 120],
  "workbench.colorTheme": "OM Theme (Default Dracula Italic)",
  "workbench.iconTheme": "material-icon-theme",
  "code-runner.executorMap": {
    "python": "clear ; python -u"
  },
  "code-runner.runInTerminal": true,
  "code-runner.ignoreSelection": true,
  "editor.fontFamily": "Consolas, 'Dank Mono', 'Source Code Pro', 'Fira Code', Menlo, 'Inconsolata', 'Droid Sans Mono', 'DejaVu Sans Mono', 'Ubuntu Mono', 'Courier New', Courier, Monaco, monospace",
  "terminal.integrated.fontFamily": "",
  "[python]": {
    "editor.defaultFormatter": "ms-python.python",
    "editor.tabSize": 4,
    "editor.insertSpaces": true,
    "editor.formatOnSave": true,
    "editor.codeActionsOnSave": {
      "source.fixAll": true,
      "source.fixAll.unusedImports": true,
      "source.organizeImports": true
    }
  },
  "python.languageServer": "Pylance",
  "python.formatting.autopep8Args": [
    "--indent-size=4",
    "--max-line-length=80"
    // "--ignore=E111"
  ],
  "python.linting.flake8Args": [
    // "--ignore=E111",
  ],
  "python.linting.flake8Enabled": true,
  "python.linting.mypyEnabled": true,
  "python.testing.unittestEnabled": false,
  "python.testing.pytestEnabled": true,
  "python.analysis.diagnosticSeverityOverrides": {},
  // "python.defaultInterpreterPath": "./venv/bin/python",
  "python.analysis.typeCheckingMode": "basic",
  "liveserver.settings.fullReload" : true, //Configuração para funcionar o Live Server > servidor de internet do python
  "cSpell.enabled": false
}


############################################################

#Settings Fabio global


{
    
    "window.zoomLevel": 2,
    "workbench.editorAssociations": [
        {
            "viewType":"jupyter.notebook.ipynb",
            "filenamePattern":"*.ipynb"
        }
    ],
    "python.linting.mypyEnabled":true,
    "python.linting.flake8Enabled": true,
    "[python]":{
        "editor.formatOnSave": true},
    "workbench.colorTheme": "Dracula",
    "files.autoSave": "afterDelay",

}

#Settings Fabio ambiente virtual (Para não ficar habilitando e desabilitando o ambiente virtual)

{ "python.pythonPath" : "C:/Fabio/Jupyter/.envs/projeto",
  "code-runner.executorMap" : {"python" : "C:/Fabio/Jupyter/.envs/projeto/bin/python",}
}
