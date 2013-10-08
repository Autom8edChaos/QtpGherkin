QtpGherkin
==========

A Gherkin language parser that will fit in QTP / VBScript. This was created as a techdemo, just to see if it is possible (and yes, it is).

This parser can understand the Gherkin keywords 'Scenario', 'Given', 'When', 'Then' and 'And'. It can recognize parameters inside the Gherkin phrases and these are extracted with a regular expression.

When using this parser, it is important that the registration of the given/when/then actions happens before the actual call to the feature runner, otherwise the actions are not recognized.

The registration of the action can be just atop of the function with the actions, this mimics the way it is done in .NET (Specflow) and Java (Cucumber). It could also be registered separately, allthough it makes less sense.
One function can be registered to multiple actions and also multiple types of actions (a Then and a When for example)

The attributes can be assigned by using:
[given] "a spaceship '(.*)'", "DefineSpaceShipUsage"
now the sentence "Given a spaceship 'SS Heart of Gold'" in the feature file will call the function DefineSpaceShipUsage(params). The array params will contain one element: params(0)="SS Heart of Gold".
