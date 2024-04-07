DocX Generator

Overview

DocX Generator generates reports in the form of word files based on template file. The template is a normal word file that is easy to edit and design. The template file is filled with data. The data is assumed to exist in Json dats format; thus allowing complex data layout to be mapped to the report.
The templated is designed by inserting directives inside label elements.

Template Directives

The directives are commands inserted into the template file to bind data, loop through data list, or customize certain behavior.
In the following, we explain each template directive:

Binding directive

The binding directive is written as @{bindingExpression}. The bindingExpression locates a certain Json memeber field. The value of the dataField addressed by the bindingExpression is copied to the output report. 
Repeat directive

The repeat directive is a scoped directive. You start the repeat scope 
with the begining directive @Repeat{listName} and close the repeat scope with the @EndRepeat directive.

Inside the repeat scope you can use the binding directive or any other directive 
to map data to the output report. Once you are inside the repeat scope, you
don't need to specify the full member name of the data field. The data field is specified relative to each list item.

Context directive

The context directive is used to simplify writing of binding expressions. 
The context directive is a scoped directive. You can start the scope by the @Context{bindingExpression} and terminate the scope by the @EndContext directive.
Inside the context directive scope, you can use binding expressions relative to the context 
binding expression. You can specify a name to the binding expression using the syntax 
@Context [contextName]{bindingExpression}.

Table directives

Tables are essential components of any report. A table presents a list of data items. A list of json data
could be mapped to rows or columns of a table. Sometimes, it is also necessary to customize the style of the table. 
We provide a powerful style control technique by allowing template writer to show/hide certain rows or columns.

Row repeat directive

The row repeat directive maps a list of json data elements to table rows. The directive is 
written as @RowRepeat[indexVariable]{bindingExpression}. The template writer should create a table 
and create a template row inside the table.
The row repeat directive can be placed inside any cell of the template row.
The binding expression should point to a Json array element. 
You specify an index variable to address the array element in binding expressions.
After placing the row repeat directive inside the template cell, you can place binding directives in other cells.
Note that the row repeat directive doesn't update the binding context. i.e. you should write binding expressions relative 
to the innermost context directive. The index variable is used to select an element from the expanded bound list.
For example, if you insert the directive @RowRepeat[i]{dataArray} inside a table template row,
you can write the binding expressions in other row cells as @{dataArray[#i].Name}.

Column repeat directive

The column repeat directive is similar to row repeat directives. It maps a list of json data elements to table columns. The directive is 
written as @ColRepeat[indexVariable]{bindingExpression}. The template writer should create a table 
and create a template column inside the table.
The column repeat directive can be placed inside any cell of the template column.
The binding expression should point to a Json array element. 
You specify an index variable to address the array element in binding expressions.
After placing the row repeat directive inside the template cell, you can place binding directives in other cells.
Note that the column repeat directive doesn't update the binding context. 
The index variable is used to select an element from the expanded bound list.

Row show/hide directives

The row show/hide directives can be used to conditionally structure a table. These directives shows/hides a
certain row based on boolean flag. The difference between them is that row show directive keeps the row visible if the
controlling flag is true while the row hide directive hides the row if the controlling flag is true. 
The row show directive is written as @RowShow{bindingExpression} while row hide directive is 
written as @RowHide{bindingExpression}. The binding expression should point to a boolean Json memeber.

Column show/hide directives

The column show/hide directives can be used to conditionally structure a table. These directives shows/hides a
certain column based on boolean flag. The difference between them is that column show directive keeps the column visible if the
controlling flag is true while the column hide directive hides the column if the controlling flag is true. 
The column show directive is written as @ColShow{bindingExpression} while col hide directive is 
written as @ColHide{bindingExpression}. The binding expression should point to a boolean Json memeber.


Binding Expressions

You can use binding expressions inside binding directive to locate a certain Json member field.
The binding expression may contain array indexing such as @{list[8]} or object member access usng the 
dot operator such as @{obj.Name}. The binding expression may also contain an index variable such as
 @{data[#i].Name}. The index variable is used in template rows or columns. The index variable #i 
 is replaced in each row by the row index.