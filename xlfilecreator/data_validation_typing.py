from typing import Dict


Header = str # 'Worker Type'
Source = str # '=Droptdownlist!$F$2:$F$3'
SourceDict = Dict[Header, Source]
            # {
                # 'Worker Gender': '=Droptdownlist!$B$2:$B$4',
                # 'Worker Pay Type Name': '=Droptdownlist!$C$2:$C$5',
                # 'Rate type': '=Droptdownlist!$F$2:$F$3'
            # }
SingleOptionsDict = Dict[Header, Dict[str,str]] 
                # {'Worker Type':{
                                # 'validate': 'list',
                                #   'source': '=Droptdownlist!$F$2:$F$3',
                                #   'error_type': 'warning',
                                #   'input_title': 'Worker Paytype',
                                #   'input_message': 'Select a value from the picklist',
                                #   'error_title': 'Input value not valid!',
                                #   'error_message': 'It should be a value from the picklist'
                                #   }
                        # }
DataValDict = Dict[Header, SingleOptionsDict]
            # {
                # 'Worker Gender': 
                        # {'validate': 'list',
                        #   'source': '=Dropdown_Lists!$A$2:$A$4',
                        #   'error_type': 'stop',
                        #   'error_title': 'Worker Gender ERROR TITLE'},
                #  'Worker Pay Type Name': 
                        # {'validate': 'list',
                        #   'source': '=Dropdown_Lists!$B$2:$B$5',
                        #   'error_type': 'information'},
                #  'ACTIVE WORKER?': 
                        # {'validate': 'list',
                        #   'source': '=Dropdown_Lists!$C$2:$C$3',
                        #   'error_type': 'stop',
                        #   'input_title': 'Active worker?',
                        #   'input_message': 'Is this worker an active employee?',
                        #   'error_message': 'it must be an option from the picklist'
            #  }

