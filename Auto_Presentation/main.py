from pyfiles.PowerpointClass import PowerPointCreator
from pyfiles.read import import_list_from_file
import imaginairy
from imaginairy.schema import ImaginePrompt

bigtitle="Main theme"
smalltitle="Subtitle"
image_on_first_page="Image_path"

#Prompt：Enter your request:teach me ________, output the answer in Json which includes at least 15 dictionaries. Each dictionary should include title, content as the dictionary key, and the content should be well-elaborated to include useful 100 words and some references.


# Create an instance of the PowerPointCreator class
ppt_creator = PowerPointCreator()
# Call the create method with the necessary arguments
ppt_creator.create(
    bigtitle,
    smalltitle,
    image_on_first_page,  # Replace with the actual path to your image
    import_list_from_file("script/Your_script"),
    "output.pptx"
)
