<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="HS_REPC_Final.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Report Cards App</title>
    <link href="Content/bootstrap.min.css" rel="stylesheet" />
    <script src="Scripts/jquery-3.3.1.min.js"></script>
    <script src="Scripts/bootstrap.min.js"></script>
    <link href="Content/flipk.css" rel="stylesheet" />
    <link href="https://fonts.googleapis.com/css?family=Roboto" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
        <div id="flipkart-navbar">
            <div class="container">
                <div class="row row1">
                    <ul class="largenav pull-right">
                        <li class="upper-links">Grades:</li>
                        <li class="upper-links"><a class="links" href="#" id="st01">PK</a></li>
                        <li class="upper-links"><a class="links" href="#" id="st0">K</a></li>
                        <li class="upper-links"><a class="links" href="#" id="st1">1</a></li>
                        <li class="upper-links"><a class="links" href="#" id="st2">2</a></li>
                        <li class="upper-links"><a class="links" href="#" id="st3">3</a></li>
                        <li class="upper-links"><a class="links" href="#" id="st4">4</a></li>
                        <li class="upper-links"><a class="links" href="#" id="st5">5</a></li>
                        <li class="upper-links"><a class="links" href="#" id="st6">6</a></li>
                        <li class="upper-links"><a class="links" id="st7" href="#">7</a></li>
                        <li class="upper-links"><a class="links" id="st8" href="#">8</a></li>
                        <li class="upper-links"><a class="links" id="st9" href="#">9</a></li>
                        <li class="upper-links"><a class="links" id="st10" href="#">10</a></li>
                        <li class="upper-links"><a class="links" id="st11" href="#">11</a></li>
                        <li class="upper-links"><a class="links" id="st12" href="#">12</a></li>
                      
                        <li class="upper-links">
                            <input class="flipkart-navbar-input " type="text" id="stnum" placeholder="Search for Products, Brands and more" name="" />
                        </li>
                        <li class="upper-links">
                               <select id="stCode" class="form-control">
                                   <option value="0">--Select Storecode--</option>
                                   <option value="1">T1</option>
                                   <option value="2">T2</option>
                                   <option value="3">T3</option>
                                   <option value="4">T1,T2,T3</option>
                               </select>
                        </li>
                        <li class="upper-links">
                               <select id="AllHR" class="form-control"><option value="0">--Select HR--</option></select>
                        </li>

                        <li class="dropdown upper-links">
                            <button type="button" class="btn btn-success">Reports</button>
                            <ul class="dropdown-menu">
                                <li ><a class="links repo5" href="#" id="btnESRep">ES Report Card</a></li>
                                <li ><a class="links repo6" href="#" id="btnMSexp">MS 6th Progress Report</a></li>
                                <li ><a class="links repo7" href="#" id="btn6RC">MS 6th Report Card</a></li>
                                <li><a class="links repo78" id="btnMSQ1" href="#">MS 7/8 Q1 Progress Report</a></li>
                                <li>
                                    <hr />
                                </li>
                                <li><a class="links repohsq1" id="btnst" href="#">HS Q1 Progress Report</a></li>
                                <li><a class="links reposhy1" id="btnstnum" href="#">HS Y1 StoredGrade Report</a></li>
                            </ul>
                        </li>
                    </ul>
                </div>

            </div>
        </div>
        <div style="margin-top:20px;">
            
            <table id="sttable" class="table table-bordered"   >
                <thead>
                    <tr class="tcolor">
                        <th style="width:5px;" >No.</th>
                        <th style="width:10px;"><input type="checkbox"  name="chkall" id="chkall"/></th>
                        <th style="width:400px;" id="stcant">NAME</th>
                        <th style="width:50px;">GRADE</th>
                        <th style="width:50px;">NUMBER</th>
                        <th style="width:400px;">HR / ADVISORY / BOHIO</th>
                    </tr>
                </thead>
                <tbody id="stbody" >
                    <tr>
                        <td>
                           
                        </td>
                        <td>..Not values yet!
                        </td>
                        <td>..Not values yet!
                        </td>
                        <td>..Not values yet!
                        </td>
                         <td>..Not values yet!
                        </td>
                        <td>..Not values yet!
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
        
    </form>
    <script type="text/javascript">
        $(document).ready(function () {
            $(".repo5,.repo6,.repo78,.repohsq1,.reposhy1").hide();
            var grad = "";
            var gra;
            $("#AllHR").hide();
            $("#AllHR").on('change',function () {
                if ($("#AllHR").val()!=0){
                var thr = $("#AllHR").val();
                hrgrade(gra, thr);
                }
            });

            $("#st01").click(function (e) {
                e.preventDefault();
                grad = "PK";
                gra = -1;
                stgrade(-1);
                hrSelect(-1);
                $("#AllHR").show();
                $(".repo6,.repo7").hide();
                $(".repo5").show();
                $(".repo78").hide();
                $(".repohsq1,.reposhy1").hide();
            })

            $("#st0").click(function (e) {
                e.preventDefault();
                grad = "K";
                gra = 0;
                stgrade(0);
                hrSelect(0);
                $("#AllHR").show();
                $(".repo6,.repo7").hide();
                $(".repo5").show();
                $(".repo78").hide();
                $(".repohsq1,.reposhy1").hide();
            })

            $("#st1").click(function (e) {
                e.preventDefault();
                grad = "1";
                gra = 1;
                stgrade(1);
                hrSelect(1);
                $("#AllHR").show();
                $(".repo6,.repo7").hide();
                $(".repo5").show();
                $(".repo78").hide();
                $(".repohsq1,.reposhy1").hide();
            })
            $("#st2").click(function (e) {
                e.preventDefault();
                grad = "2";
                gra = 2;
                stgrade(2);
                hrSelect(2);
                $("#AllHR").show();
                $(".repo6,.repo7").hide();
                $(".repo5").show();
                $(".repo78").hide();
                $(".repohsq1,.reposhy1").hide();
            })
            $("#st3").click(function (e) {
                e.preventDefault();
                grad = "3";
                gra = 3;
                stgrade(3);
                hrSelect(3);
                $("#AllHR").show();
                $(".repo6,.repo7").hide();
                $(".repo5").show();
                $(".repo78").hide();
                $(".repohsq1,.reposhy1").hide();
            })
            $("#st4").click(function (e) {
                e.preventDefault();
                grad = "4";
                gra = 4;
                stgrade(4);
                hrSelect(4);
                $("#AllHR").show();
                $(".repo6,.repo7").hide();
                $(".repo5").show();
                $(".repo78").hide();
                $(".repohsq1,.reposhy1").hide();
            })

            $("#st5").click(function (e) {
                e.preventDefault();
                grad = "5";
                gra = 5;
                stgrade(5);
                hrSelect(5);
                $("#AllHR").show();
                $(".repo6,.repo7").hide();
                $(".repo5").show();
                $(".repo78").hide();
                
                $(".repohsq1,.reposhy1").hide();
            })

            $("#st6").click(function (e) {
                e.preventDefault();
                stgrade(6);
                $("#AllHR").hide();
                $(".repo6,.repo7").show();
                $(".repo5").hide();
                $(".repo78").hide();
                $(".repohsq1,.reposhy1").hide();
            })
            $("#st7").click(function (e) {
                e.preventDefault();
                stgrade(7);
                $("#AllHR").hide();
                $(".repo6,.repo7").hide();
                $(".repo5").hide();
                $(".repo78").show();
                $(".repohsq1,.reposhy1").hide();

                $(".repohsq1,.reposhy1").hide();
            })
            $("#st8").click(function (e) {
                e.preventDefault();
                stgrade(8);
                $("#AllHR").hide();
                $(".repo78").show();
                $(".repohsq1,.reposhy1").hide();
                $(".repo6,.repo7").hide();
                $(".repo5").hide();
            })
            $("#st9").click(function (e) {
                e.preventDefault();
                stgrade(9);
                $("#AllHR").hide();
                $(".repohsq1,.reposhy1").show();
                $(".repo6,.repo7").hide();
                $(".repo5").hide();

            })
            $("#st10").click(function (e) {
                e.preventDefault();
                stgrade(10);
                $("#AllHR").hide();
                $(".repohsq1,.reposhy1").show();
                $(".repo6,.repo7").hide();
                $(".repo5").hide();
            })
            $("#st11").click(function (e) {
                e.preventDefault();
                stgrade(11);
                $("#AllHR").hide();
                $(".repohsq1,.reposhy1").show();
                $(".repo6,.repo7").hide();
                $(".repo5").hide();
            })
            $("#st12").click(function (e) {
                e.preventDefault();
                stgrade(12);
                $("#AllHR").hide();
                $(".repohsq1,.reposhy1").show();
                $(".repo6,.repo7").hide();
                $(".repo5").hide();
            })

            

            $("#chkall").click(function () {
                if ($(this).prop("checked")){
                    chkall();
                    $("#stnum").val(selectSch().substring(selectSch(), selectSch().length - 1));
                } else {
                    unchkall();
                    $("#stnum").val('');
                }
            });


            $("#btnstnum").click(function () {
                if ($("#stnum").val() != '') {
                    reportCard($("#stnum").val());
                }
            });
            $("#btnst").click(function () {
                if ($("#stnum").val() != '') {
                    HSQ1($("#stnum").val());
                }
            });

            
            $("#btnESRep").click(function () {
                if ($("#stnum").val() != '') {
                    ES5REPORTCARD($("#stnum").val(), grad);
                }
            });

            $("#btnMSQ1").click(function () {
                if ($("#stnum").val() != '') {
                    PROGREPORTQ1($("#stnum").val());
                }
            });
            $("#btnMSexp").click(function () {
                if ($("#stnum").val() != '') {
                    EXP_REPORTQ1($("#stnum").val());
                }
            });
            $("#btn6RC").click(function () {
                if ($("#stnum").val() != '') {
                    MSXIXREPORCARD($("#stnum").val());
                }
            });
            
            $("#stbody").on('click', '.chk', function () {
                if ($(this).prop("checked")) { 
                $("#stnum").val(selectSch().substring(selectSch(),selectSch().length-1));
                } else {
                    $("#stnum").val(selectSch().substring(selectSch(), selectSch().length - 1));
                }
            });

         
            

        });

        function hrSelect(gra){

            var param = "'gra':'" + gra + "'";
            $.ajax({
                type: "POST",
                async: false,
                url: "Default.aspx/hrSelect",
                data: "{" + param + "}",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    var m = response.d;
                    var data = '';
                    var con = 0;
                    data = '<option value="0" >--Select Teacher--</option>';
                    if (m != '') {
                        var b = m.split('^');
                        for (var i = 0; i < b.length - 1; i++) {
                            var c = b[i].split('|');
                            data += '<option Value="'+c[0]+'" >'+c[1]+'</option>';
                        }
                         document.getElementById("AllHR").innerHTML = data;
                        
                    } else {
                        document.getElementById("AllHR").innerHTML = data;
                       
                    }

                    return false;
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    alert("Respuesta = " + XMLHttpRequest.responseText + "\n Estatus = " + textStatus + "\n Error = " + errorThrown);
                }
            });
        }

        function selectSch() {
            var actSelec = "";
            $('#stbody').find(':checkbox').each(function () {
                if ($(this).prop("checked")) {
                    actSelec += this.id + ';';
                }

            });
            return actSelec;
        }

        function unchkall() {
            
            $('#stbody').find(':checkbox').each(function () {
                $(this).prop("checked", false);
            });
            
        }

        function chkall() {
            $('#stbody').find(':checkbox').each(function () {
                $(this).prop("checked", true);
            });
        }

        function hrgrade(grd, thom) {
              var param = "'gra':'" + grd + "','Htea':" + thom + "";
            $.ajax({
                type: "POST",
                async: false,
                url: "Default.aspx/hrgrade",
                data: "{" + param + "}",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    var m = response.d;
                    var data = '';
                    var con = 0;
                    if (m != '') {

                        var b = m.split('^');
                        for (var i = 0; i < b.length - 1; i++) {
                            var c = b[i].split('|');
                            con += 1;
                            data += '<tr class="chktr" ><td style="text-align:left">' + con + '</td>';
                            data += '<td> <input type="checkbox" class="chk" id="' + c[0] + '" name="chkstnum" /></td>';
                            data += '<td style="text-align:left">' + c[1] + '</td>';
                            data += '<td style="text-align:center">' + c[2] + '</td>';
                            data += '<td style="text-align:center">' + c[0] + '</td>';
                            data += '<td style="text-align:left">' + c[3] + '</td>';
                            data += '</tr>';

                        }

                        document.getElementById("stbody").innerHTML = data;
                        document.getElementById("stcant").innerHTML = 'NAME (' + con + ')';

                    } else {
                        data += '<tr><td colspan="5">No data</td></tr>';
                        document.getElementById("stbody").innerHTML = data;
                        document.getElementById("stcant").innerHTML = 'NAME (' + con + ')';

                    }

                    return false;
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    alert("Respuesta = " + XMLHttpRequest.responseText + "\n Estatus = " + textStatus + "\n Error = " + errorThrown);
                }
            });
        }
        function stgrade(grd) {
            
            var param = "'gra':'" + grd + "'";
            $.ajax({
                type: "POST",
                async: false,
                url: "Default.aspx/stgrade",
                data: "{" + param + "}",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    var m = response.d;
                    var data = '';
                    var con = 0;
                    if (m != '') {
                       
                        var b = m.split('^');
                        for (var i = 0; i < b.length - 1; i++) {
                            var c = b[i].split('|');
                            con += 1;
                            data += '<tr class="chktr" ><td style="text-align:left">' + con + '</td>';
                            data += '<td> <input type="checkbox" class="chk" id="' + c[0] + '" name="chkstnum" /></td>';
                            data += '<td style="text-align:left">' + c[1] + '</td>';
                            data += '<td style="text-align:center">' + c[2] + '</td>';
                            data += '<td style="text-align:center">' + c[0] + '</td>';
                            data += '<td style="text-align:left">' + c[3] + '</td>';
                            data += '</tr>';
                            
                        }
                        
                        document.getElementById("stbody").innerHTML = data;
                        document.getElementById("stcant").innerHTML = 'NAME (' + con + ')';
                        
                    } else {
                        data += '<tr><td colspan="5">No data</td></tr>';
                        document.getElementById("stbody").innerHTML = data;
                        document.getElementById("stcant").innerHTML = 'NAME (' + con + ')';
                       
                    }

                    return false;
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    alert("Respuesta = " + XMLHttpRequest.responseText + "\n Estatus = " + textStatus + "\n Error = " + errorThrown);
                }
            });
        }

        function ES5REPORTCARD(stid,gr) {
            var param = "'stnum':'" + stid + "','grade':'" + gr + "'";

            $.ajax({
                type: "POST",
                async: false,
                url: "Default.aspx/ES5REPORTCARD",
                data: "{" + param + "}",
               contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    var m = response.d;
                    if (m != '') {
                        window.open("RepoFiles/" + m);
                    } else {
                        alert("Student with not Historical Grades!");
                    }
                    return false;
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    alert("Respuesta = " + XMLHttpRequest.responseText + "\n Estatus = " + textStatus + "\n Error = " + errorThrown);
                }
            });
        }

        
        function MSXIXREPORCARD(stid) {
            var param = "'stnum':'" + stid + "'";

            $.ajax({
                type: "POST",
                async: false,
                url: "Default.aspx/MSXIXREPORCARD",
                data: "{" + param + "}",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    var m = response.d;
                    if (m != '') {
                        window.open("RepoFiles/" + m);
                    } else {
                        alert("Student with not Historical Grades!");
                    }
                    return false;
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    alert("Respuesta = " + XMLHttpRequest.responseText + "\n Estatus = " + textStatus + "\n Error = " + errorThrown);
                }
            });
        }

        function EXP_REPORTQ1(stid) {
            var param = "'stnum':'" + stid + "'";

            $.ajax({
                type: "POST",
                async: false,
                url: "Default.aspx/EXP_REPORTQ1",
                data: "{" + param + "}",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    var m = response.d;
                    if (m != '') {
                        window.open("RepoFiles/" + m);
                    } else {
                        alert("Student with not Historical Grades!");
                    }
                    return false;
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    alert("Respuesta = " + XMLHttpRequest.responseText + "\n Estatus = " + textStatus + "\n Error = " + errorThrown);
                }
            });
        }

        function PROGREPORTQ1(stid) {
            var param = "'stnum':'" + stid + "'";

            $.ajax({
                type: "POST",
                async: false,
                url: "Default.aspx/PROGREPORTQ1",
                data: "{" + param + "}",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    var m = response.d;
                    if (m != '') {
                        window.open("RepoFiles/" + m);
                    } else {
                        alert("Student with not Historical Grades!");
                    }
                    return false;
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    alert("Respuesta = " + XMLHttpRequest.responseText + "\n Estatus = " + textStatus + "\n Error = " + errorThrown);
                }
            });
        }

        function HSQ1(stid) {
            var param = "'stnum':'" + stid + "'";

            $.ajax({
                type: "POST",
                async: false,
                url: "Default.aspx/HSQ1",
                data: "{" + param + "}",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    var m = response.d;
                    if (m != '') {
                        window.open("RepoFiles/" + m);
                    } else {
                        alert("Student with not Historical Grades!");
                    }
                    return false;
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    alert("Respuesta = " + XMLHttpRequest.responseText + "\n Estatus = " + textStatus + "\n Error = " + errorThrown);
                }
            });
        }

        function reportCard(stid) {
            var param = "'stnum':'" + stid + "'";

            $.ajax({
                type: "POST",
                async: false,
                url: "Default.aspx/reportCard",
                data: "{" + param + "}",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    var m = response.d;
                    if (m != '') {
                        window.open("RepoFiles/" + m);
                    } else {
                        alert("Student with not Historical Grades!");
                    }
                    return false;
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    alert("Respuesta = " + XMLHttpRequest.responseText + "\n Estatus = " + textStatus + "\n Error = " + errorThrown);
                }
            });
        }
    
    </script>
</body>
</html>
