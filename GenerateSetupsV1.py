import adsk.core, adsk.fusion, adsk.cam, traceback
ui = None
app= None
cam = None
commandId = 'GenerateSetupsCommand'
commandName = 'Generate Setups'
commandDescription = 'Generate setups from a template for each sheet in a nesting study'

handlers = []

def run(context):
    try:
        # Get the application and userInterface
        global app
        app = adsk.core.Application.get()
        global ui
        ui = app.userInterface

        # Get the active product.
        global cam
        cam = adsk.cam.CAM.cast(app.activeProduct)

        # Check if anything was returned.
        if (cam is None):
            ui.messageBox('The Manufacturing workspace must be active')
            return
        
        global commandId
        global commandName
        global commandDescription
        
        # Create the command definition and open the inputs panel
        cmdDef = ui.commandDefinitions.itemById(commandId)
        if not cmdDef:
            cmdDef = ui.commandDefinitions.addButtonDefinition(commandId, commandName, commandDescription) # no resource folder is specified, the default one will be used

        onCommandCreated = GenerateSetupsCreatedHandler()
        cmdDef.commandCreated.add(onCommandCreated)
        # keep the handler referenced beyond this function
        handlers.append(onCommandCreated)

        inputs = adsk.core.NamedValues.create()
        cmdDef.execute(inputs)

        # prevent this module from being terminate when the script returns, because we are waiting for event handlers to fire
        adsk.autoTerminate(False)
        
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
    
# Gets a list of templates from the Cloud Library        
def getTemplates():
    libManager = adsk.cam.CAMManager.get().libraryManager
    templateLibrary = libManager.templateLibrary
    templates = templateLibrary.childTemplates(templateLibrary.urlByLocation(adsk.cam.LibraryLocations.CloudLibraryLocation))
    return templates
            
# Creates setups for each nesting sheet and applies the selected template
def generateSetups(template):
    setups: adsk.cam.Setups = cam.setups
        
    for model in cam.manufacturingModels:
        # TODO - handle QTY > 1
        setupInput = setups.createInput(adsk.cam.OperationTypes.MillingOperation)
        setupInput.models=[model.occurrence]
        setupInput.stockMode=adsk.cam.SetupStockModes.FixedBoxStock
        setupParams = setupInput.parameters
        setupParams.itemByName('job_stockFixedX').value.value=120.0
        setupParams.itemByName('job_stockFixedXOffset').value.value=0.0
        setupParams.itemByName('job_stockFixedY').value.value=240.0
        setupParams.itemByName('job_stockFixedYOffset').value.value=0.0
        setupParams.itemByName('job_stockFixedZ').value.value=1.8
        setupParams.itemByName('job_stockFixedZOffset').value.value=0.0
        setupParams.itemByName('job_groundStockModelOrigin').value.value=True
        setupParams.itemByName('wcs_orientation_mode').value.value="modelOrientation"
        setupParams.itemByName('wcs_origin_mode').value.value="modelOrigin"
        
        setups.add(setupInput)
        
    for setup in setups:
        setup.createFromCAMTemplate(template)
        
        
        pocketOp: adsk.cam.Operation = None
        for op in setup.operations:
            
            if op.name.startswith('POCKET'):
                pocketOp = op
                break
        if (pocketOp is None):
            ui.messageBox('No Pocket operation in template')
            return
        
        faces = getPocketFaces(setup.models)
        if len(faces) > 0:
            pocketGeometryParam: adsk.cam.CadContours2dParameterValue = pocketOp.parameters.itemByName('pockets').value
            selections = pocketGeometryParam.getCurveSelections()
            selections.clear()
            selection = selections.createNewPocketSelection()
            selection.inputGeometry=faces
            pocketGeometryParam.applyCurveSelections(selections)
            cam.generateToolpath(pocketOp)
        else:
            pocketOp.deleteMe()

#finds all faces that have all vertices with z between 8mm and 10mm
def getPocketFaces(manufacturingModels):
    pocketFaces = []
    for model in manufacturingModels:
        components = model.childOccurrences
        for component in components:
            bodies= component.bRepBodies
            for body in bodies:
                faces = body.faces
                for face in faces:
                    vertices = face.vertices
                    pocket = True
                    for vertex in vertices:
                        # geometry is in cm
                        if vertex.geometry.z < 0.8 or vertex.geometry.z >1.0:
                            pocket = False
                            break;
                    if pocket:
                        pocketFaces.append(face)
    return pocketFaces

# Everything from here down is mostly boilerplate code for creating the options dialog
class GenerateSetupsExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            cmd = args.firingEvent.sender
            inputs = cmd.commandInputs
            
            templateListInput = None
            for inputI in inputs:
                global commandId
                if inputI.id == commandId + '_templateList':
                    templateListInput = inputI
           
            templates = getTemplates()
            template = None
            for t in templates:
                if t.name == templateListInput.selectedItem.name:
                    template = t
                    break
            if not template:
                if ui:
                    ui.messageBox('Template not found.')
                return
            
            generateSetups(template)

        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

class GenerateSetupsDestroyHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            # when the command is done, terminate the script
            # this will release all globals which will remove all event handlers
            adsk.terminate()
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

class GenerateSetupsCreatedHandler(adsk.core.CommandCreatedEventHandler):    
    def __init__(self):
        super().__init__()        
    def notify(self, args):
        try:
            cmd: adsk.core.Command = args.command
            cmd.isRepeatable = False
            onExecute = GenerateSetupsExecuteHandler()
            cmd.execute.add(onExecute)
            
            onDestroy = GenerateSetupsDestroyHandler()
            cmd.destroy.add(onDestroy)
            
            # keep the handler referenced beyond this function
            handlers.append(onExecute)
            handlers.append(onDestroy)
            
            inputs = cmd.commandInputs
            global commandId
           
            templateListInput = inputs.addDropDownCommandInput(commandId + '_templateList', 'Template', adsk.core.DropDownStyles.TextListDropDownStyle)
            # appearances = getAppearancesFromLib(materialLibNames[0], '')
            templates = getTemplates()
            listItems = templateListInput.listItems
            for template in templates:
                if template.name == 'Wikihouse Blocks v10':
                    listItems.add(template.name, True, '')
                else:
                    listItems.add(template.name, False, '')
            
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))