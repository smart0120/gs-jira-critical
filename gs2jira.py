#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script for converting google sheet rows to Jira tickets
"""

import os, gspread, time
from gspread.exceptions import GSpreadException
from dotenv import load_dotenv
from jira import JIRA
from jira.exceptions import JIRAError
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta
from datetime import date, timedelta

__author__ = "bursno22"
__license__ = "MIT"
__version__ = "0.0.1"

load_dotenv()

def index_from_col(col_name):
    """
    Return index from column name
    For example, 0 for A, 1 for B, etc
    """
    return ord(col_name.upper()) - 65

def main():
    # Open Google Sheet
    gc = gspread.oauth()
    sh = gc.open(os.getenv('SHEET_NAME'))
    primary_worksheet = sh.get_worksheet(int(os.getenv('PRIMARY_SHEET')))
    jira_server_url = os.getenv('JIRA_SERVER_URL')

    critical_system_item_list = primary_worksheet.col_values(index_from_col(os.getenv('ITEM_NAME')))
    tool_owner_list = primary_worksheet.col_values(index_from_col(os.getenv('TOOL_OWNER')))

    row_range = [int(val) for val in os.getenv('DATA_RANGE').split(':')]

    for row in range(row_range[0], row_range[1]+1):
        item_name = critical_system_item_list[row]
        tool_owner = tool_owner_list[row]

        template = {
            "type": "doc",
            "version": 1,
            "content": [{
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "Application: ",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    },
                    {
                        "text": item_name,
                        "type": "text"
                    },
                ]
            },
            {
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "Business Owner: ",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    },
                    {
                        "text": tool_owner,
                        "type": "text"
                    },
                ]
            },
            {
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "Data Owner: ",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    },
                    {
                        "text": tool_owner,
                        "type": "text"
                    },
                    {
                        "type": "hardBreak"
                    }
                ]
            },
            {
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "Overview",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    }
                ]
            },
            {
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "IT controls are established to ensure that particular requirements driven by internal policies, procedures, standards or by regulatory requirements are in place and effective."
                    }
                ]
            },
            {
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "Moreover, the IT Controls are "
                    },
                    {
                        "type": "text",
                        "text": "required by regulations",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "text": " such as "
                    },
                    {
                        "type": "text",
                        "text": "BalT from BaFin",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "text": "  (the regulatory authority that provides our Banking license)."
                    }
                ]
            },
            {
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "As such, we require your complete engagement to ensure the successful execution of our planned controls for this year.",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    },
                    {
                        "type": "hardBreak"
                    }
                ]
            },
            {
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "Next Steps",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    }
                ]
            },
            {
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "1. "
                    },
                    {
                        "type": "text",
                        "text": "Review",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "text": " the IT Controls applicable to system below, noting "
                    },
                    {
                        "type": "text",
                        "text": "key dates",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "text": " and incorporating them into your team’s 2021 roadmap."
                    }
                ]
            },
            {
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "2. "
                    },
                    {
                        "type": "text",
                        "text": "Nominate a delegate",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "text": " from your team who will be engaged to execute the IT Control (tag their name in the ″Nominated Delegate″ Column)."
                    }
                ]
            },
            {
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "3. "
                    },
                    {
                        "type": "text",
                        "text": "Flag any concerns",
                        "marks": [
                            {
                                "type": "strong"
                            }
                        ]
                    },
                    {
                        "type": "text",
                        "text": " you have in the comments of this ticket (e.g. timeline conflicts, unclear IT Control guidelines, etc.)"
                    },
                    {
                        "type": "hardBreak"
                    }
                ]
            },
            {
                "type": "paragraph",
                "content": [
                    {
                        "type": "text",
                        "text": "For more details in the 2021 IT Controls see "
                    },
                    {
                        "type": "text",
                        "text": os.getenv('DOC_URL'),
                        "marks": [
                            {
                                "type": "link",
                                "attrs": {
                                    "href": os.getenv('DOC_URL')
                                },
                            }
                        ]
                    },
                    {
                        "type": "hardBreak"
                    }
                ]
            },
            {
                "type": "table",
                "attrs": {
                    "isNumberColumnEnabled": False,
                    "layout": "default"
                },
                "content": [
                    {
                        "type": "tableRow",
                        "content": [
                            {
                                "type": "tableHeader",
                                "content": [
                                    {
                                        "type": "paragraph",
                                        "content": [
                                            {
                                                "type": "text",
                                                "text": "IT Control",
                                                "marks": [
                                                    {
                                                        "type": "strong"
                                                    }
                                                ]
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "tableHeader",
                                "content": [
                                    {
                                        "type": "paragraph",
                                        "content": [
                                            {
                                                "type": "text",
                                                "text": "Target Date",
                                                "marks": [
                                                    {
                                                        "type": "strong"
                                                    }
                                                ]
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "tableHeader",
                                "content": [
                                    {
                                        "type": "paragraph",
                                        "content": [
                                            {
                                                "type": "text",
                                                "text": "Nominated Delegate",
                                                "marks": [
                                                    {
                                                        "type": "strong"
                                                    }
                                                ]
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "tableHeader",
                                "content": [
                                    {
                                        "type": "paragraph",
                                        "content": [
                                            {
                                                "type": "text",
                                                "text": "JIRA Ticket",
                                                "marks": [
                                                    {
                                                        "type": "strong"
                                                    }
                                                ]
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "tableHeader",
                                "content": [
                                    {
                                        "type": "paragraph",
                                        "content": [
                                            {
                                                "type": "text",
                                                "text": "Oversight Team",
                                                "marks": [
                                                    {
                                                        "type": "strong"
                                                    }
                                                ]
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    }
                ]
            }]
        }

        issue_dict = {
            'project': os.getenv('JIRA_PROJECT_KEY'),
            'summary': f'{item_name} 2021 IT Control Action Plan',
            'description': template,
            'issuetype': {'name': os.getenv('JIRA_TICKET_TYPE')}
        }

        try:
            # Open JIRA
            auth_jira = JIRA(
                options={'server': jira_server_url, 'rest_api_version': 3},
                basic_auth=(os.getenv('JIRA_USERNAME'), os.getenv('JIRA_OAUTH_TOKEN'))
            )
            issue_key = str(auth_jira.create_issue(fields=issue_dict))
            print(f'create new ticket {issue_key}')
            
        except JIRAError as err:
            print(str(err))

if __name__ == '__main__':
    main()
