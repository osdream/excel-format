== the sheet meta ==
{{{sheet.name}}}

== use excel axis ==
{{{1,A}}} {{{2,B}}} !

== use number indexes ==
Yes, {{{3,3}}} {{{4,2}}} {{{5,1}}}

== complex case ==
<% for (var i = 'A'.charCodeAt(0); i <= 'E'.charCodeAt(0); i++) { %>
{{{7,<%= String.fromCharCode(i) %>}}}
<% } %>
