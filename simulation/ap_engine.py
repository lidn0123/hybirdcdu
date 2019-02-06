"""
Performs the aspen simulation

Attributes:
    aspen                   COM object which links Python with Aspen Plus
"""
import os
import win32com.client as win32

project_path = os.path.abspath('../')
model_folder = 'simulation\hbcdumodel'
model_name = 'CDU-basic.apw'

aspen = win32.Dispatch('Apwn.Document')
print("--- loading aspen .bkp file ---")
print('Simulation path: ', os.path.join(project_path, model_folder, model_name))
aspen.InitFromArchive2(os.path.join(project_path, model_folder, model_name))
print("--- .bkp file loaded---")
aspen.Visible = True
print('If file open falied, please open it manually')
print('File open? Y/N?')
file_statue = input()
if file_statue in ['Y', 'y']:
    aspen.Activate()

else:
    print('---Close Aspen File---')
    aspen.Close()
    aspen.Quit()
    del aspen


def run(temperature, pressure, buoh_flow_in, acac_flow_in):
    """
    Runs the Aspen Plus simulatin placed on 'simulation/PBR.bkp' and returns a descriptor object
    :param temperature: in Celsius
    :param pressure: in bar
    :param buoh_flow_in: molar flow of buthanol in kmol/h
    :param acac_flow_in: molar flow of acetic acid in kmol/h
    :return: a descriptor object with the simulation inputs and results
    """
    # print("--- setting inputs ---")
    # response = PbrStatus(temperature, pressure, buoh_flow_in, acac_flow_in)
    aspen.Tree.FindNode("\Data\Streams\FEED\Input\TEMP\MIXED").Value = temperature
    aspen.Tree.FindNode("\Data\Blocks\PBR\Input\PRES").Value = pressure
    aspen.Tree.FindNode("\Data\Streams\FEED\Input\FLOW\MIXED\BUTANOL").Value = buoh_flow_in
    aspen.Tree.FindNode("\Data\Streams\FEED\Input\FLOW\MIXED\ACACETIC").Value = acac_flow_in

    # print("\n--- runnning ---")
    try:
        print("\n--- runnning simulation---")
        aspen.Engine.Run2()
        print("--- simulation finished---")
        catalyst_weight = aspen.Application.Tree.FindNode(
            "\Data\Flowsheeting Options\Design-Spec\PBRWT\Output\FINAL_VAL\\1").Value
        buoh_flow_out = aspen.Tree.FindNode("\Data\Streams\OUTPUT\Output\MOLEFLOW\MIXED\BUTANOL").Value
        water_flow_out = aspen.Tree.FindNode("\Data\Streams\OUTPUT\Output\MOLEFLOW\MIXED\AGUA").Value
        acac_flow_out = aspen.Tree.FindNode("\Data\Streams\OUTPUT\Output\MOLEFLOW\MIXED\ACACETIC").Value
        buac_flow_out = aspen.Tree.FindNode("\Data\Streams\OUTPUT\Output\MOLEFLOW\MIXED\BUT-ACET").Value
        vap_fraction = aspen.Tree.FindNode("\Data\Streams\OUTPUT\Output\VFRAC_OUT\MIXED").Value
        temperature_out = aspen.Tree.FindNode("\Data\Streams\OUTPUT\Output\TEMP_OUT\MIXED").Value
        # response.update_result(catalyst_weight, buoh_flow_out, water_flow_out, acac_flow_out, buac_flow_out, vap_fraction, temperature_out)
    except:
        print("--- ASPEN COM ERROR ---\nFailed with")
        print("temperature : %s" % temperature)
        print("pressure : %s" % pressure)
        print("buoh_flow_in : %s" % buoh_flow_in)
        print("acac_flow_in : %s" % acac_flow_in)

    # return response


def close_aspen():
    print("--- closing aspen simulation file ---")
    aspen.close()
