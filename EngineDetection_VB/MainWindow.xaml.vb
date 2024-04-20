Class MainWindow
    'window opens: only option is to choose report type (Event or Player)
    '''Event path
    '''expose input box for event name
    '''validate input against database. if no events of that name are found, error out
    '''expose input box for game source (values from dim.Source.SourceName). if the event entered only has one possible source, pre-populate it and do not allow field value to change
    '''if there's multiple options, give user option to choose which one

    '''Player path
    '''expose input boxes for player first and last names
    '''validate inputs against database. if no players are found, error out
    '''expose input box for game source (values from dim.Source). if the event entered only has one possible source, pre-populate it and do not allow field value to change
    '''if there's multiple options, give user option to choose which one
    '''expose date entries for start and end dates and allow user to choose the dates. validate to ensure start date is on or before the end date

    'for both paths, guide user in selecting the comparison dataset. All options should come from the DB and/or be predefined
    '''1. choose a source
    '''2. choose a time control
    '''3. choose a ratingID
    '''4. choose a score name

    'other variables:
    '''engine - get from DB instead of an input parameter
    '''depth - get from DB instead of an input parameter
    '''max eval - seems like getting from the DB would be better
End Class
