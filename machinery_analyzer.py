import pandas as pd
import numpy as np
import re

class MachineryAnalyzer:
    def __init__(self):
        """Initialize the MachineryAnalyzer with critical machinery list and patterns."""
        self.critical_machinery = [
            'Main Engine', 'Auxiliary Engine', 'Steering Gear', 'Main Switchboard',
            'Emergency Generator', 'Navigation Equipment', 'Communication Equipment',
            'Boiler', 'Main Air Compressor', 'Main Cooling Water Pump',
            'Fire Pump', 'Bilge Pump', 'Oily Water Separator'
        ]

        self.trim_suffix_patterns = [
            # ✅ NEW: handles suffixes stuck directly to machinery names (no space/hyphen)
            r'(Fwd|Aft|Port|Starboard|Fwd-Port|Aft-Port|Fwd-Stbd|Aft-Stbd)$',

            # ✅ Existing logic for direction suffixes
            r'(?:Fwd|Aft|Port|Starboard|Fwd-Port|Aft-Port|Fwd-Stbd|Aft-Stbd)[\s\-]*$',

            # ✅ e.g. " 20 Person"
            r'\s+\d+\s*Person$',

            # ✅ e.g. "No", "No6"
            r'(?:No\d*|No)$',

            # ✅ e.g. "#2", "3"
            r'#?\d+$',

            # ✅ e.g. " Port", "-Starboard"
            r'[\s\-]+(Port|Starboard)$'
        ]

        self.update_values = {
            'AnchorP ': 'Anchor',
            'AnchorPort1 ': 'Anchor',
            'AnchorStarboard1': 'Anchor',
            'Main Engine - ME-C II': 'Main Engine',
            'Main Engine - MC': 'Main Engine',
            'Anchor Chain CableP ': 'Anchor Chain Cable',
            'Anchor Chain CablePort1': 'Anchor Chain Cable',
            'Anchor Chain CableStarboard1 ': 'Anchor Chain Cable',
            'Accommodation LadderPort1': 'Accommodation Ladder',
            'Lifeboat Davit Hydraulic UnitAft1 ' : 'Lifeboat Davit Hydraulic Unit',
            'Pilot LadderPort1': 'Pilot Ladder',
            'Accommodation LadderStarboard1': 'Accommodation Ladder',
            'LiferaftStarboard2 ' : 'Liferaft',
            'LiferaftPort2 ': 'Liferaft',
            'Chain LockerStarboard1 ': 'Chain Locker',
            'Chain LockerPort1 ': 'Chain Locker',
            'Accommodation LadderS ': 'Accommodation Ladder',
            'Void SpacesNoBELOW AFT  ACCOMMODATION GANGWAY ': 'Void Spaces',
            'Void SpacesNoC DECK ': 'Void Spaces',
            'Void SpacesNoBRIDGE DECK ': 'Void Spaces',
            'Void SpacesNo6P':'Void Spaces',
            'Void SpacesNo7S': 'Void Spaces',
            'Bunker DavitPort1 ': 'Bunker Davit',
            'LiferaftStarboard1' : 'Liferaft',
            'Anchor Chain CableS ': 'Anchor Chain Cable',
            'Bunker DavitP ': 'Bunker Davit',
            'AnchorS ': 'Anchor',
            'Combined Windlass Mooring WinchP ': 'Combined Windlass Mooring Winch',
            'Combined Windlass Mooring WinchS ': 'Combined Windlass Mooring Winch',
            'Chain LockerS ': 'Chain Locker',
            'Chain LockerP ': 'Chain Locker',
            'Bunker DavitS ': 'Bunker Davit',
            'Lifeboat DavitAft1': 'Lifeboat Davit',
            'Lifeboat Davit Hydraulic UnitAft1': 'Lifeboat Davit Hydraulic Unit',
            'LifeboatAft1 ': 'Lifeboat',
            'Provision CranePort1': 'Provision Crane',
            'Accommodation LadderP ': 'Accommodation Ladder',
            'LiferaftStarboard2': 'Liferaft',
            'Chain LockerStarboard1': 'Chain Locker',
            'LiferaftPort2': 'Liferaft',
            'Chain LockerPort1': 'Chain Locker',
            'AnchorPort1': 'Anchor',
            'Anchor Chain CableStarboard1': 'Anchor Chain Cable',
            'Bunker DavitPort1': 'Bunker Davit',
            'Bunker DavitStarboard1': 'Bunker Davit',
            'LifeboatAft1': 'Lifeboat',
            'Pilot LadderStarboard1': 'Pilot Ladder',
            'Pilot Combination LadderStarboard1': 'Pilot Combination Ladder',
            'Void SpacesNo7P': 'Void Spaces',
            'Provision CraneStarboard1': 'Provision Crane',
            'Void SpacesNo6S': 'Void Spaces',
            'LiferaftForward1': 'Liferaft',
            'Void SpacesNo4P': 'Void Spaces',
            'LiferaftPort1': 'Liferaft',
            'Void SpacesNo2P': 'Void Spaces',
            'Rescue BoatStarboard1': 'Rescue Boat',
            'Void SpacesNoFORE PEAK': 'Void Spaces',
            'Void SpacesNo2S': 'Void Spaces',
            'Void SpacesNo4S': 'Void Spaces',
            'Pilot Combination LadderPort1': 'Pilot Combination Ladder',
            'Bilge WellCentre1': 'Bilge Well',
            'Bilge WellStarboard1': 'Bilge Well',
            'Bilge WellPort1 ':'Bilge WellPort1',
            'Combined Mooring Winch Hydraulic UnitForward1': 'Combined Mooring Winch Hydraulic Unit',
            'Mooring Winch Hydraulic UnitAft1': 'Mooring Winch Hydraulic Unit',
            'Combined Windlass Mooring WinchForward2': 'Combined Windlass Mooring Winch',
            'Combined Windlass Mooring WinchForward1': 'Combined Windlass Mooring Winch',
            'Air HornAft1': 'Air Horn',
            'Bilge WellPort1': 'Bilge Well',
            'Liferaft 6 PersonNo1': 'Liferaft',
            'LiferaftNo4': 'Liferaft',
            'Liferaft 16 PersonNo2': 'Liferaft',
            'LiferaftNo5':'Liferaft',
            'Liferaft/Rescue Boat DavitStarboard1': 'Liferaft/Rescue Boat Davit',
            'LifeboatPort1': 'Lifeboat',
            'RefrigeratorNo1': 'Refrigerator',
            'RefrigeratorNo2' : 'Refrigerator',
            'Lifeboat DavitPort1': 'Lifeboat Davit',
            'Liferaft 16 PersonNo3': 'Liferaft',
            'Mooring WinchStarboard2': 'Mooring Winch',
            'Combined Windlass Mooring WinchStarboard ':'Combined Windlass Mooring Winch',
            'Combined Windlass Mooring WinchPort ': 'Combined Windlass Mooring Winch',
            'Mooring WinchStarboard3':'Mooring Winch',
            'Mooring WinchPort ': 'Mooring Winch',
            'Pilot LadderPort ': 'Pilot Ladder',
            'Pilot LadderStarboard ':'Pilot Ladder',
            'Emergency Towing SystemAft ':'Emergency Towing System',
            'Emergency Towing SystemForward ':'Emergency Towing System',
            'Mooring WinchStarboard1':'Mooring Winch',
            'Anchor Chain CablePort ':'Anchor Chain Cable',
            'Accommodation LadderPort ': 'Accommodation Ladder',
            'Anchor Chain CableStarboard ':'Anchor Chain Cable',
            'AnchorPort ': 'Anchor',
            'Chain LockerPort': 'Chain Locker',
            'AnchorStarboard': 'Anchor',
            'Provision CraneStarboard ':'Provision Crane',
            'Chain LockerPort ':'Chain Locker',
            'AnchorStarboard ': 'Anchor',
            'Main Engine - RT - A': 'Main Engine',
            'Pilot Combination LadderPort' :'Pilot Combination Ladder',
            'Lifeboat/Rescue BoatPort ':'Lifeboat/Rescue Boat',
            'Machine Parts Handling CraneStarboard ':'Machine Parts Handling Crane',
            'Liferaft Embarkation LadderPort ':'Liferaft Embarkation Ladder',
            'LifeboatStarboard ':'Lifeboat',
            'Liferaft Embarkation LadderStarboard ':'Liferaft Embarkation Ladder',
            'Liferaft Embarkation LadderForward ':'Liferaft Embarkation Ladder',
            'Chain LockerStarboard':'Chain Locker',
            'Lifeboat DavitStarboard ':'Lifeboat Davit',
            'Accommodation LadderStarboard ':'Accommodation Ladder',
            'Machine Parts Handling Crane':'Machine Parts Handling Crane',
            'Pilot Combination LadderStarboard' : 'Pilot Combination Ladder',
            'Rescue Boat DavitPort ':'Rescue Boat Davit',
            'Bunker DavitStarboard':'Bunker Davit',
            'Bunker DavitStarboard ':'Bunker Davit',
            'Pilot Combination LadderPort ':'Pilot Combination Ladder',
            'Bunker DavitPort ':'Bunker Davit',
            'Machine Parts Handling CranePort ':'Machine Parts Handling Crane',
            'Pilot Combination LadderStarboard ':'Pilot Combination Ladder',
            'Liferaft Embarkation LadderStarboard1':'Liferaft Embarkation Ladder',
            'LiferaftStarboard ':'Liferaft',
            'Combined Windlass Mooring WinchFwd-Stbd1':'Combined Windlass Mooring Winch',
            'Liferaft Embarkation LadderPort1':'Liferaft Embarkation Ladder',
            'Liferaft Embarkation LadderPort2':'Liferaft Embarkation Ladder',
            'Liferaft Embarkation LadderStarboard2':'Liferaft Embarkation Ladder',
            'Mooring WinchCentre1':'Mooring Winch',
            'Combined Windlass Mooring WinchFwd-Port1':'Combined Windlass Mooring Winch',
            'Liferaft/Rescue Boat DavitStarboard ': 'Liferaft/Rescue Boat Davit',
            'Rescue BoatStarboard ':'Rescue Boat',
            'Mooring Winch Hydraulic UnitAft ':'Mooring Winch Hydraulic Unit',
            'Miscellaneous Handling DavitPort ':'Miscellaneous Handling Davit',
            'Chain LockerStarboard ':'Chain Locker',
            'LiferaftPort ':'Liferaft',
            'Main Engine - ME-C II GI': 'Main Engine',
            'Lifeboat DavitAft ':'Lifeboat Davit',
            'Liferaft/Rescue Boat DavitPort ':'Liferaft/Rescue Boat Davit',
            'LifeboatAft ':'Lifeboat',
            'Auxiliary EngineNo2':'Auxiliary Engine',
            'Auxiliary EngineNo1':'Auxiliary Engine',
            'Auxiliary EngineNo3':'Auxiliary Engine',
            'Auxiliary EngineNo4':'Auxiliary Engine',
            'Combined Windlass Mooring WinchPort1':'Combined Windlass Mooring Winch',
            'Mooring WinchPort1':'Mooring Winch',
            'Liferaft 25 PersonPort1':'Liferaft',
            'Liferaft 25 PersonStarboard1':'Liferaft',
            'Combined Windlass Mooring WinchStarboard1':'Combined Windlass Mooring Winch',
            'Provision CranePort': 'Provision Crane',
            'Cargo Hose Handling CranePort ':'Cargo Hose Handling Crane',
            'ICCPAft ':'ICCP',
            'Cargo Hose Handling CraneStarboard ':'Cargo Hose Handling Crane',
            'Combined Mooring Winch Hydraulic UnitForward ':'Combined Mooring Winch Hydraulic Unit',
            'Mooring WinchFwd-Port1':'Mooring Winch',
            'Clear View ScreenPort1':'Clear View Screen',
            'Mooring WinchForward1':'Mooring Winch',
            'Miscellaneous Handling DavitStarboard1':'Miscellaneous Handling Davit',
            'Mooring WinchAft1':'Mooring Winch',
            'Mooring WinchAft2':'Mooring Winch',
            'Suez Search Light DavitForward1':'Suez Search Light Davit',
            'LiferaftForward ':'Liferaft',
            'LifeboatAft':'Lifeboat',
            'Mooring WinchAft3':'Mooring Winch',
            'Mooring WinchAft-Port1':'Mooring Winch',
            'Clear View ScreenStarboard1':'Clear View Screen',
            'Liferaft Embarkation LadderNoAFT - STBD':'Liferaft Embarkation Ladder',
            'Liferaft Embarkation LadderNoFWD - STBD':'Liferaft Embarkation Ladder',
            'Liferaft Embarkation LadderNoAFT - PORT ':'Liferaft Embarkation Ladder',
            'Liferaft Embarkation LadderNoFWD - PORT':'Liferaft Embarkation Ladder',
            'Void SpacesNo1C':'Void Spaces',
            'Void SpacesNoFORE PEAK 1C':'Void Spaces',
            'Auxiliary EngineNo5':'Auxiliary Engine',
            'Mooring WinchAft-Stbd1 ' : 'Mooring Winch',
            'Void SpacesNo8P':'Void Spaces',
            'LiferaftNoFWD':'Liferaft',
            'Mooring WinchFwd-Stbd1':'Mooring Winch',
            'Mooring WinchAft-Port2':'Mooring Winch',
            'Lifeboat/Rescue BoatStarboard1':'Lifeboat/Rescue Boat',
            'Void SpacesNoAFT PEAK 1C':'Void Spaces',
            'Void SpacesNo8S':'Void Spaces',
            'Lifeboat Davit.Starboard1':'Lifeboat Davit',
            'Mooring WinchAft-Stbd1':'Mooring Winch',
            'LiferaftNoFWD ':'Liferaft',
            'Mooring WinchMiddle2':'Mooring Winch',
            'Mooring WinchMiddle1':'Mooring Winch',
            'Provision CraneStarboard':'Provision Crane',
            'Pilot LadderPort':'Pilot Ladder',
            'Pilot LadderStarboard':'Pilot Ladder',
            'Liferaft 15 PersonPort2':'Liferaft',
            'Liferaft 15 PersonPort1':'Liferaft',
            'Combined Windlass Mooring WinchPort':'Combined Windlass Mooring Winch',
            'Lifeboat DavitAft':'Lifeboat Davit',
            'Sludge DavitPort ':'Sludge Davit',
            'Combined Windlass Mooring WinchStarboard':'Combined Windlass Mooring Winch',
            'Sludge DavitStarboard':'Sludge Davit',
            'Sludge DavitPort':'Sludge Davit',
            'Sludge DavitStarboard ':'Sludge Davit',
            'Liferaft/Rescue Boat DavitStarboard':'Liferaft/Rescue Boat Davit',
            'Bunker DavitPort':'Bunker Davit',
            'Mooring WinchCentre2':'Mooring Winch',
            'Provision Crane SPort1':'Provision Crane',
            'LiferaftFwd-Stbd1':'Liferaft',
            'EPIRBStarboard1':'EPIRB',
            'Mooring WinchAft':'Mooring Winch',
            'Combined Windlass Mooring WinchForward':'Combined Windlass Mooring Winch',
            'Mooring WinchForward':'Mooring Winch',
            'Muster StationAft':'Muster Station',
            'Suez Search Light DavitForward':'Suez Search Light Davit',
            'SARTStarboard':'SART',
            'Combined Windlass Mooring WinchForward ':'Combined Windlass Mooring Winch',
            'Suez Search Light DavitForward ':'Suez Search Light Davit',
            'SARTStarboard ':'SART',
            'Muster StationAft ':'Muster Station',
            'Mooring Winch Hydraulic UnitAft':'Mooring Winch Hydraulic Unit',
            'Mooring Winch Hydraulic UnitAft ':'Mooring Winch Hydraulic Unit',
            'Combined Mooring Winch Hydraulic UnitForward':'Combined Mooring Winch Hydraulic Unit',
            'Combined Mooring Winch Hydraulic UnitForward ':'Combined Mooring Winch Hydraulic Unit',
            'Mooring WinchForward':'Mooring Winch',
            'Mooring WinchForward ':'Mooring Winch',
            'SARTPort':'SART',
            'SARTPort ':'SART',
            'Mooring WinchAft':'Mooring Winch',
            'Mooring WinchAft ':'Mooring Winch',
            'Muster StationAft':'Muster Station',
            'Rescue BoatPort':'Rescue Boat',
            'Rescue BoatPort ':'Rescue Boat',
            'Rescue BoatPort  ':'Rescue Boat',
            'Provision CranePort':'Provision Crane',
            'Provision CranePort ':'Provision Crane',
            'Provision CranePort  ':'Provision Crane',
            'VLSFO Transfer PumpForward1':'VLSFO Transfer Pump',
            'ICCPForward':'ICCP',
            'VLSFO Transfer PumpAft':'VLSFO Transfer Pump',
            'Main CSW PumpNo3':'Main CSW Pump',
            'VLSFO Transfer PumpForward2':'VLSFO Transfer Pump',
            'Main CSW PumpNo2':'Main CSW Pump',
            'Main CSW PumpNo1':'Main CSW Pump',
            'ICCPForward ':'ICCP',
            'Mooring Winch Hydraulic UnitForward':'Mooring Winch Hydraulic Unit',
            'Mooring Winch Hydraulic UnitForward ':'Mooring Winch Hydraulic Unit',
            'Liferaft 25 Person':'Liferaft',
            'Liferaft 25 Person ':'Liferaft',
            'Main Engine - WinGD':'Main Engine',
            'Mooring Winch Hydraulic UnitForward':'Mooring Winch Hydraulic Unit',
            'Combined Windlass Mooring WinchAft':'Combined Windlass Mooring Winch',
            'Accommodation LadderPort':'Accommodation Ladder',
            'Combined Windlass Mooring WinchAft':'Combined Windlass Mooring Winch',
            'AnchorPort':'Anchor',
            'Accommodation LadderStarboard':'Accommodation Ladder',
            'Combined Windlass Mooring WinchAft ':'Combined Windlass Mooring Winch',
            'Liferaft 6 Person':'Liferaft',
            'Main Engine - RTFLEX': 'Main Engine',
            'Main Engine - RTFLEX # 1':'Main Engine',
            'Main Engine - RTFLEX#1':'Main Engine',
            'Mooring WinchStarboard':'Mooring Winch',
            'Main Engine - ME - B':'Main Engine',
            'Mooring WinchStarboard':'Mooring Winch',
            'LiferaftPort':'Liferaft',
            'Mooring WinchPort':'Mooring Winch',
            'Anchor Chain CableStarboard':'Anchor Chain Cable',
            'LiferaftStarboard':'Liferaft',
            'Anchor Chain CablePort':'Anchor Chain Cable',
            'Mooring Winchstarboard':'Mooring Winch',
            'Mooring Winchstarboard ':'Mooring Winch',
            'Mooring Winchstarboard  ':'Mooring Winch',
            'Pilot Ladder Davitport1':'Pilot Ladder Davit',
            'Pilot Ladder Davitstarboard1':'Pilot Ladder Davit',
            'Mooring Winchport2':'Mooring Winch',
            'Pilot Ladder Davitstarboard1 ':'Pilot Ladder Davit',
            'Pilot Ladder Davitport1 ':'Pilot Ladder Davit',
            'Pilot Ladder Davitstarboard1  ':'Pilot Ladder Davit',
            'Pilot Ladder Davitport1  ':'Pilot Ladder Davit',
            'Liferaftfwd-Stbd':'Liferaft',
            'Suez Search Light Davitcentre':'Suez Search Light Davit',
            'Liferaft Embarkation Ladderfwd-Port':'Liferaft Embarkation Ladder',
            'Muster Stationstarboard':'Muster Station',
            'Liferaftfwd-Stbd ':'Liferaft',
            'Suez Search Light Davitcentre ':'Suez Search Light Davit',
            'Liferaft Embarkation Ladderfwd-Port ':'Liferaft Embarkation Ladder',
            'Muster Stationstarboard ':'Muster Station',
            'Liferaft Embarkation Ladderstarboard':'Liferaft Embarkation Ladder',
            'Liferaft Embarkation Ladderport':'Liferaft Embarkation Ladder',
            'Clear View ScreenPort':'Clear View Screen',
            'Clear View ScreenStarboard':'Clear View Screen',
            'Rescue BoatStarboard':'Rescue Boat',
            'Liferaft Embarkation LadderStarboard':'Liferaft Embarkation Ladder',
            'Muster StationStarboard':'Muster Station',
            'Liferaft Embarkation LadderPort':'Liferaft Embarkation Ladder',
            'Machine Parts Handling CranePort':'Machine Parts Handling Crane',
            'Muster StationPort':'Muster Station',
            'Liferaft Embarkation LadderFwd-Port':'Liferaft Embarkation Ladder',
            'Liferaft Embarkation Ladderforward':'Liferaft Embarkation Ladder',
            'Machine Parts Handling CraneStarboard':'Machine Parts Handling Crane',
            'Suez Search Light DavitCentre':'Suez Search Light Davit',
            'LiferaftFwd-Stbd':'Liferaft',
            'Liferaft/Rescue Boat DavitPort':'Liferaft/Rescue Boat Davit',
            'LifeboatStarboard':'Lifeboat',
            'Liferaft Embarkation LadderForward':'Liferaft Embarkation Ladder',
            'Lifeboat/Rescue BoatPort':'Lifeboat/Rescue Boat',
            'Lifeboat DavitStarboard':'Lifeboat Davit',
            'Rescue Boat DavitPort':'Rescue Boat Davit',
            'Provision CranePort2':'Provision Crane',
            'Lifeboat/Rescue Boat DavitStarboard1':'Lifeboat/Rescue Boat Davit',
            'Provision CranePort3':'Provision Crane',
            'Lifeboat DavitPort':'Lifeboat Davit',
            'LifeboatPort' :'Lifeboat',
            'Bilge WellAft':'Bilge Well',
            'Bilge WellPort' :'Bilge Well',
            'Bilge WellStarboard' :'Bilge Well',
            'LiferaftForward':'Liferaft',
            'Cargo Gear Room Hatch CoverForward1':'Cargo Gear Room Hatch Cover',
            'Cargo Gear Room Hatch CoverAft1':'Cargo Gear Room Hatch Cover',
            'Portable Gas DetectorNo2':'Portable Gas Detector',
            'Portable Gas DetectorNo3':'Portable Gas Detector',
            'Portable Gas DetectorNo4':'Portable Gas Detector',
            'Personal Gas DetectorNo5':'Personal Gas Detector',
            'Personal Gas DetectorNo1':'Personal Gas Detector',
            'Liferaft Embarkation LadderForward1':'Liferaft Embarkation Ladder',
            'Liferaft Embarkation LadderForward2':'Liferaft Embarkation Ladder',
            'Accommodation LadderStarboard2':'Accommodation Ladder',
            'Mooring WinchAft4':'Mooring Winch',
            'Mooring WinchForward2':'Mooring Winch',
            'Liferaft 6 PersonForward':'Liferaft',
            'ICCPAft':'ICCP',
            'Liferaft Embarkation LadderFwd-Stbd':'Liferaft Embarkation Ladder',
            'SCBA No1 Bridge':'SCBA',
            'Rescue BoatPort1':'Rescue Boat',
            "Fireman'S Outfit No1 Fire Locker - Bridge":"Fireman'S Outfit",
            'EPIRBStarboard':'EPIRB',
            'Liferaft Embarkation LadderFwd-Stbd':'Liferaft Embarkation Ladder',
            'SCBA No3 Mid/Ship Deck Store':'SCBA',
            'SCBA No4 Mid/Ship Deck Store':'SCBA',
            'Mooring WinchAft-Stbd2':'Mooring Winch',
            'Mooring WinchAft-Stbd4':'Mooring Winch',
            'Bilge WellStarboard5':'Bilge Well',
            'Bilge WellPort3':'Bilge Well',
            'Bilge WellPort7':'Bilge Well',
            'Bilge WellPort2':'Bilge Well',
            'Mooring WinchAft-Port3':'Mooring Winch',
            'Liferaft 16 PersonPort1':'Liferaft',
            'Liferaft 15 PersonStarboard2':'Liferaft',
            'Bilge WellStarboard6':'Bilge Well',
            'Liferaft 16 PersonPort2':'Liferaft',
            'Sewage Discharge PumpPort':'Sewage Discharge Pump',
            'Bilge WellStarboard7':'Bilge Well',
            'Bilge WellPort4':'Bilge Well',
            'Bilge WellPort5':'Bilge Well',
            'Bilge WellPort6':'Bilge Well',
            'Liferaft 15 PersonStarboard1':'Liferaft',
            'Bilge WellStarboard4':'Bilge Well',
            'Bilge WellStarboard3':'Bilge Well',
            'Bilge WellStarboard2':'Bilge Well',
            'Bilge WellPort 5':'Bilge Well',
            'Mooring Winch Hydraulic UnitAft2':'Mooring Winch Hydraulic Unit',
            'Combined Mooring Winch Hydraulic UnitForward2':'Combined Mooring Winch Hydraulic Unit',
            'Lifeboat/Rescue BoatStarboard':'Lifeboat/Rescue Boat',
            'Miscellaneous Handling DavitPort':'Miscellaneous Handling Davit',
            'Pilot Ladder DavitStarboard':'Pilot Ladder Davit',
            'Ballast PumpNo2':'Ballast Pump',
            'Mooring WinchAft-Stbd':'Mooring Winch',
            'LifeboatStarboard1':'Lifeboat',
            'Lifeboat/Rescue BoatPort1':'Lifeboat/Rescue Boat',
            'Lifeboat DavitStarboard1':'Lifeboat Davit',
            'Monorail CraneStarboard1':'Monorail Crane',
            'Ballast PumpNo1':'Ballast Pump',
            'Monorail CranePort1':'Monorail Crane',
            'Combined Windlass Mooring WinchFwd-Port':'Combined Windlass Mooring Winch',
            'Bilge EductorForward1':'Bilge Eductor',
            'Liferaft 35 Person':'Liferaft',
            'Pilot Combination LadderPort2':'Pilot Combination Ladder',
            'Liferaft 6 PersonForward1':'Liferaft',
            'Cargo Pump-Cargo Tank No5 P':'Cargo Pump-Cargo Tank',
            'Cargo Pump-Cargo Tank No6 S':'Cargo Pump-Cargo Tank',
            'Cargo Pump-Cargo Tank No2 P':'Cargo Pump-Cargo Tank',
            'Fixed Tank Cleaning Machinery-Cargo Tank No6':'Fixed Tank Cleaning Machinery-Cargo Tank',
            'Cargo Pump-Cargo Tank No 1 S':'Cargo Pump-Cargo Tank',
            'Cargo Pump-Cargo Tank No2 S':'Cargo Pump-Cargo Tank',
            'Fixed Tank Cleaning Machinery-Cargo Tank No7':'Fixed Tank Cleaning Machinery-Cargo Tank',
            'Emergency Towing SystemForward':'Emergency Towing System',
            'Emergency Towing SystemAft':'Emergency Towing System',
            'Cargo Pump-Cargo Tank No6 P':'Cargo Pump-Cargo Tank',
            'Cargo Pump-Cargo Tank No8 S':'Cargo Pump-Cargo Tank',
            'Fixed Tank Cleaning Machinery-Cargo Tank No4':'Fixed Tank Cleaning Machinery-Cargo Tank',
            'Cargo Pump-Cargo Tank No4 P':'Cargo Pump-Cargo Tank',
            'Cargo Pump-Cargo Tank No 4 S':'Cargo Pump-Cargo Tank',
            'Fixed Tank Cleaning Machinery-Cargo Tank No9':'Fixed Tank Cleaning Machinery-Cargo Tank',
            'Cargo Pump- Slope Tank P':'Cargo Pump- Slope Tank',
            'Cargo Pump-Cargo Tank No3 P':'Cargo Pump-Cargo Tank',
            'Fixed Tank Cleaning Machinery-Cargo Tank No2':'Fixed Tank Cleaning Machinery-Cargo Tank',
            'Cargo Pump-Cargo Tank No1 P':'Cargo Pump-Cargo Tank',
            'Cargo Pump-Cargo Tank No5 S':'Cargo Pump-Cargo Tank',
            'Cargo Pump-Slope Tank S':'Cargo Pump-Slope Tank',
            'Cargo Pump-Cargo Tank No3 S':'Cargo Pump-Cargo Tank',
            'Cargo Pump-Cargo Tank No9 S':'Cargo Pump-Cargo Tank',
            'Bridge Wing Control ConsoleStarboard':'Bridge Wing Control Console',
            'Combined Windlass Mooring WinchFwd-Stbd2':'Combined Windlass Mooring Winch',
            'Cargo Pump-Cargo Tank No8 P':'Cargo Pump-Cargo Tank',
            'Fixed Tank Cleaning Machinery-Cargo Tank No3':'Fixed Tank Cleaning Machinery-Cargo Tank',
            'Fixed Tank Cleaning Machinery-Cargo Tank No8':'Fixed Tank Cleaning Machinery-Cargo Tank',
            'Cargo Pump-Cargo Tank No7 S':'Cargo Pump-Cargo Tank',
            'Cargo Pump-Cargo Tank No9 P':'Cargo Pump-Cargo Tank',
            'Fixed Tank Cleaning Machinery-Cargo Tank No5':'Fixed Tank Cleaning Machinery-Cargo Tank',
            'Fixed Tank Cleaning Machinery-Cargo Tank No1':'Fixed Tank Cleaning Machinery-Cargo Tank',
            'Cargo Pump-Cargo Tank No 7 P':'Cargo Pump-Cargo Tank',
            'Bridge Wing Control ConsolePort':'Bridge Wing Control Console',
            'Cargo Pump-Cargo Tank No 1 S':'Cargo Pump-Cargo Tank',
            'Cargo Pump-Cargo Tank No4 P':'Cargo Pump-Cargo Tank',
            'Cargo Pump-Cargo Tank No1 S':'Cargo Pump-Cargo Tank',
            'Cargo Deep Well PumpPort4':'Cargo Deep Well Pump',
            'Cargo Deep Well PumpPort1' :'Cargo Deep Well Pump',
            'Cargo Deep Well PumpPort2' :'Cargo Deep Well Pump',
            'Cargo Deep Well PumpPort3' :'Cargo Deep Well Pump',
            'Cargo Deep Well PumpPort5' :'Cargo Deep Well Pump',
            'Cargo Deep Well PumpPort6' :'Cargo Deep Well Pump',
            'Cargo Deep Well PumpStarboard2':'Cargo Deep Well Pump',
            'Cargo Deep Well PumpStarboard1':'Cargo Deep Well Pump',
            'Cargo Deep Well PumpStarboard3':'Cargo Deep Well Pump',
            'Cargo Deep Well PumpStarboard4':'Cargo Deep Well Pump',
            'Cargo Deep Well PumpStarboard5':'Cargo Deep Well Pump',
            'Cargo Deep Well PumpStarboard6':'Cargo Deep Well Pump',
            'Lifeboat Davit Hydraulic UnitPort1':'Lifeboat Davit Hydraulic Unit',
            'Cargo Pump - Slop TankPort1':'Cargo Pump - Slop Tank',
            'Lighting TransformerForward1':'Lighting Transformer',
            'Ballast Water Treatment PlantAft1':'Ballast Water Treatment Plant',
            'Ballast Water Treatment PlantStarboard1':'Ballast Water Treatment Plant',
            'Liferaft/Rescue Boat DavitPort1':'Liferaft/Rescue Boat Davit',
            'Ballast Water Treatment PlantPort1':'Ballast Water Treatment Plant',
            'Cargo Pump - Slop TankStarboard1':'Cargo Pump - Slop Tank',
            'Cargo Pump - Residue TankPort1':'Cargo Pump - Residue Tank',
            'Mooring WinchFwd-Stbd': 'Mooring Winch',
            'Liferaft 20 PersonPort':'Liferaft',
            'Miscellaneous Handling DavitPort':'Miscellaneous Handling Davit',
            'Liferaft 20 PersonStarboard' :'Liferaft',
            'Mooring WinchFwd-Port':'Mooring Winch',
            'Mooring WinchAft-Stbd':'Mooring Winch',
            'Miscellaneous Handling CraneStarboard':'Miscellaneous Handling Crane'


        }

        

    # Position words PMS appends to machinery names, longest first so "Fwd-Port" wins over "Port"
    POSITIONS = r'Fwd-Port|Fwd-Stbd|Aft-Port|Aft-Stbd|Starboard|Forward|Centre|Center|Middle|Port|Stbd|Fwd|Aft'
    # Glued to the name in lower case only for unambiguous words ("Winchstarboard"), never "port"/"aft"
    # which also end real words ("Liferaft", "Shaft")
    LOWER_GLUED_POSITIONS = r'fwd-port|fwd-stbd|aft-port|aft-stbd|starboard|forward'

    def _lookup_update_values(self, name):
        key = re.sub(r'\s+', ' ', name).strip().lower()
        if not hasattr(self, '_update_lookup'):
            self._update_lookup = {re.sub(r'\s+', ' ', k).strip().lower(): v for k, v in self.update_values.items()}
        return self._update_lookup.get(key)

    def clean_machinery_location(self, machinery_name):
        """Machinery name without unit numbers and position, e.g. "LiferaftPort1" -> "Liferaft".

        Pass the name in its original case: a position glued to the name is recognised by its
        capital letter ("LifeboatAft"), which keeps "Liferaft" and "Shaft" intact.
        """
        if not isinstance(machinery_name, str):
            return machinery_name

        name = re.sub(r'\s+', ' ', machinery_name.strip())
        mapped = self._lookup_update_values(name)
        if mapped is not None:
            return mapped

        # Unit number once only, so model numbers stay ("Inmarsat-F77#1" -> "Inmarsat-F77")
        name = re.sub(r'\s*(?:#|No\.?)?\s*\d+$', '', name)                         # "#2", "No2", " 3"
        previous = None
        while name != previous:
            previous = name
            name = re.sub(r'(?:(?<=[a-z])No|\sNo\.?)$', '', name)                   # "PumpNo"
            name = re.sub(r'\s+\d+\s*Person$', '', name, flags=re.IGNORECASE)       # " 20 Person"
            name = re.sub(rf'[\s\-]+(?:{self.POSITIONS})$', '', name, flags=re.IGNORECASE)  # " Port", "-Aft"
            name = re.sub(rf'(?<=[a-z0-9.)/])(?:{self.POSITIONS})$', '', name)      # "LifeboatAft"
            name = re.sub(rf'(?<=[a-z])(?:{self.LOWER_GLUED_POSITIONS})$', '', name)  # "Winchstarboard"
            name = name.strip(' -')

        mapped = self._lookup_update_values(name)
        if mapped is not None:
            return mapped
        # "Main Engine - WX", "Main Engine - ME-C II GI", ... are all the main engine itself
        # (but not "Main Engine - HCU" or "Main Engine - Safety and Control")
        if re.fullmatch(r'Main Engine\s*-\s*(?:ME|MC|RT|UEC|WX?|X|WinGD|DF|GI)\b[\w\s\-\.]*', name, flags=re.IGNORECASE):
            return 'Main Engine'
        return name


    def is_critical(self, machinery_name):
        if not isinstance(machinery_name, str):
            return False
        cleaned_name = self.clean_machinery_location(machinery_name)
        return any(critical.lower() in cleaned_name.lower() for critical in self.critical_machinery)

    def cleaned_set(self, names):
        """Lower-case cleaned machinery names, for comparing vessel and reference lists."""
        cleaned = pd.Series(names).dropna().astype(str).map(self.clean_machinery_location)
        return {str(x).strip().lower() for x in cleaned if str(x).strip()}

    def process_data(self, data, ref_sheet):
        try:
            # Same reference lists as the QuickView summary, so both report the same missing machinery
            sheets = pd.read_excel(ref_sheet, sheet_name=None)
            reference_data = sheets.get('Machinery Location', next(iter(sheets.values())))
            critical_data = sheets.get('Critical Machinery', pd.DataFrame())
            vessel_specific = sheets.get('Vessel Specific Machinery', pd.DataFrame())

            if 'Machinery Location' not in data.columns or 'Machinery Location' not in reference_data.columns:
                return data, {"error": "Required columns missing in data or reference sheet"}

            # Work on a copy so the shared vessel data keeps its original columns
            data = data.copy()
            data['Machinery Location Clean'] = data['Machinery Location'].map(self.clean_machinery_location)
            data['Critical'] = data['Machinery Location'].apply(self.is_critical)

            vml_set = self.cleaned_set(data['Machinery Location'])
            ml_set = self.cleaned_set(reference_data['Machinery Location'])
            cm_set = self.cleaned_set(critical_data.get('Critical Machinery', pd.Series(dtype=str)))
            vsm_set = self.cleaned_set(vessel_specific.get('Vessel Specific Machinery', pd.Series(dtype=str)))

            # Combine all known expected machinery
            union_set = ml_set.union(cm_set, vsm_set)

            # Compare
            difMachineryVessel = vml_set - union_set
            missingmachinery = union_set - vml_set

            # Format output
            difMachineryVessel_df = pd.DataFrame({'Different Machinery on Vessel': sorted(difMachineryVessel)})
            difMachineryVessel_df['Different Machinery on Vessel'] = difMachineryVessel_df['Different Machinery on Vessel'].str.title()

            missingmachinery_df = pd.DataFrame({'Missing Machinery on Vessel': sorted(missingmachinery)})
            missingmachinery_df['Missing Machinery on Vessel'] = missingmachinery_df['Missing Machinery on Vessel'].str.title()

            result = {
                'different_machinery': difMachineryVessel_df,
                'missing_machinery': missingmachinery_df,
            }

            return data, result

        except Exception as e:
            print(f"Error processing machinery data: {str(e)}")
            return data, {"error": str(e)}
