# Contributing
First off, thank you for considering contributing to the project. The tools I've built have been mostly for my personal use so they can be rough.

## Submitting Changes
### Versioning
I tend to follow [Semanitic Versioning] (https://semver.org/) as introduced by [Hadley Wickham](http://r-pkgs.had.co.nz/description.html#version) using three (3) number separated by a period.
 + MAJOR: Generally includes significant new features and/or breaking changes to the code.
 + MINOR: Generally includes new features, but can also be a push of a number of bug fixes (or a mix). Only backward compatible feautres and fixes can be included here, with few exceptions (i.e. deprication without removal).
 + PATCH: Fixing a bug without adding new features. Only backwards compatible fixes may be included here.
 
For not-yet-stable tools or those undergoing significant initial core development, a four (4) digit code following 0.0.0.9000 should be used and only the fourth (4th) digit should be incremented. At 1.0.0, the core tool should be stable and generally usable.

*All suggestions are welcome. Naming conventions, or whitespace, or more performant code--doesn't matter.*

**This section is in development as tools enter different stages.**
### Major Release Submissions
Submit a pull request. For major release changes, I expect to leave the topic open for discussion and testing for some time before making the change. 

### Minor Release Submissions
Submit a pull request. For minor release changes, I expect to leave the topic open long enough to have a healthy discussion, but expect them to be added somewhat quickly.

### Patch Submissions
Submit a pull request. For patches (and suggested improvements), if nobody can give a reason why it is a bad move, I'll likely make the change right away.

## Coding Conventions
Coding conventions, what are those? I would like to say that these tools have a consistent set of conventions, but they don't. Again, I've generally built these for my personal use and left a lot to be desired in this area.

### R and Python
#### Both
+ Be liberal with comments (hastags everywhere!)
+ camelCase for variables and underscores for functions (myVariable vs. my_function())
+ Be explicit when using functions ( ggplot(data =d at) not ggplot(dat))
+ I'm not particular on indents vs spaces, as long as it is legible.
+ Limit nesting of functions.
+ Provide spacing around arguments (x = 5 not x=5, ggplot(data = dat) not ggplot(data=dat))

#### R Specific
+ Tidyverse syntax preferred

#### Python Specific
Nothing here yet

### Power BI
#### DAX
+ Capitalize functions (e.g. SUM() not sum(), VAR not var)

#### Power Query
+ Limit nesting table functions.

### PowerShell
Nothing here yet
