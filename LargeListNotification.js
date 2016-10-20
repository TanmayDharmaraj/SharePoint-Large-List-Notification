/*!
 * Author: Tanmay Darmaraj
 * Last Modified Date: 2016-10-20
 * Revision: 1
 * Description: Large List  */

"use strict";
define(['jquery'], function(jQuery) {
    var RC = RC || {}
    RC.Messages = (function() {
        return {ErrorViewCreation: "There was a problem creating the view.", ErrorRetrievingValues: "Error retrieving values.", ErrorInsufficientPermission: "You do not have enough permissions to perform this action."}
    })();
    RC.Constants = [
        {
            locale: 1033,
            error: "This view cannot be displayed because it exceeds the list view threshold (5000 items) enforced by the administrator."
        }
    ]
    RC.Config = [
        {
            type: 0,
            view_name: "Most Recent",
            query: "<OrderBy><FieldRef Name='ID' Ascending='FALSE' /></OrderBy>"
        }, {
            type: 1,
            view_name: "Last Week",
            query: "<Where><Geq><FieldRef Name='Modified' /><Value Type='DateTime'><Today OffsetDays='-7' /></Value></Geq></Where>"
        }, {
            type: 2,
            view_name: "Last 30 days",
            query: "<Where><Geq><FieldRef Name='Modified' /><Value Type='DateTime'><Today OffsetDays='-30' /></Value></Geq></Where>"
        }
    ];
    RC.Helpers = (function() {
        function GetView(options) {
            var deferred = jQuery.Deferred();
            jQuery.ajax({
                url: _spPageContextInfo.siteAbsoluteUrl + "/_api/web/lists('" + options.ListId + "')/views/getbytitle('" + options.ViewName + "')?" + options.OData,
                type: 'GET',
                headers: {
                    "accept": "application/json;odata=verbose"
                }
            }).done(function(data) {
                deferred.resolve(data);
            }).fail(function(err) {
                deferred.reject(err);
            })
            return deferred.promise();
        }

        function CheckPermissions(permissions_array) {
            var dfd = jQuery.Deferred();
            var ctx = new SP.ClientContext.get_current();
            var web = ctx.get_web();
            var ob = new SP.BasePermissions();
            for (var i = 0; i < permissions_array.length; i++) {
                //ob.set(SP.PermissionKind.manageWeb)
                ob.set(permissions_array[i]);
            }
            var per = web.doesUserHavePermissions(ob)
            ctx.executeQueryAsync(function() {
                var isAllowed = per.get_value();
                dfd.resolve(isAllowed)
            }, function(a, b) {
                dfd.reject(b.get_message());
            });
            return dfd.promise();
        }
        return {GetView: GetView, CheckPermissions: CheckPermissions}
    })();

    RC.Services = (function() {
        function CreateDefaultView(requested_view_type) {
            var view_configuration = null;
            for (var i = 0; i < RC.Config.length; i++) {
                if (RC.Config[i].type == requested_view_type) {
                    view_configuration = RC.Config[i];
                    break;
                }
            }
            var ctx = SP.ClientContext.get_current();
            var web = ctx.get_web();
            var listCollection = web.get_lists();
            var list = listCollection.getById(_spPageContextInfo.pageListId);
            var viewCollection = list.get_views();
            ctx.load(viewCollection);
            ctx.executeQueryAsync(function() {
                //get the default view
                var default_view = null;
                for (var index = 0; index < viewCollection.get_count(); index++) {
                    var current_view = viewCollection.itemAt(index);
                    if (current_view.get_defaultView()) {
                        default_view = current_view;
                        break;
                    }
                }
                for (var index = 0; index < viewCollection.get_count(); index++) {
                    var current_view = viewCollection.itemAt(index);
                    if (current_view.get_title() == view_configuration.view_name) {
                        //We found a view with a same name. Cannot create the view automatically.
                        var url = current_view.get_serverRelativeUrl()
                        window.location.href = window.location.protocol + "//" + window.location.host + url;
                        return;
                    }
                }
                var default_view_fields = default_view.get_viewFields()
                ctx.load(default_view_fields);
                ctx.executeQueryAsync(function() {
                    var new_view_fields = [];

                    for (var i = 0; i < default_view_fields.get_count(); i++) {
                        new_view_fields.push(default_view_fields.getItemAtIndex(i));
                    }
                    //create a new view
                    var new_view = new SP.ViewCreationInformation();
                    new_view.set_title(view_configuration.view_name);

                    var camlQuery = new SP.CamlQuery();
                    var query = view_configuration.query;
                    camlQuery.set_viewXml(query);
                    new_view.set_query(camlQuery);
                    new_view.set_viewFields(new_view_fields)
                    new_view.set_paged(true);
                    viewCollection.add(new_view);
                    ctx.load(viewCollection);
                    ctx.executeQueryAsync(function() {
                        //We have now created the view. To get the url of the view, be a good guy and don't infer but get the value from REST
                        RC.Helpers.GetView({ListId: _spPageContextInfo.pageListId, OData: "jQueryselect=ServerRelativeUrl", ViewName: view_configuration.view_name}).done(function(data) {
                            var url = data.d.ServerRelativeUrl
                            window.location.href = window.location.protocol + "//" + window.location.host + url;
                        }).fail(function(err) {
                            console.error(RC.Messages.ErrorRetrievingValues);
                            console.error(err);
                        })
                    }, function(sender, args) {
                        console.error(RC.Messages.ErrorViewCreation);
                        console.error(args.get_message());
                    });
                }, function(sender, args) {
                    console.error(RC.Messags.ErrorRetrievingValues);
                    console.error(args.get_message());
                })

            }, function(sender, args) {
                console.error(args.get_message());
            });
        }

        function CreatePersonalView(requested_view_type) {
            var view_configuration = null;
            for (var i = 0; i < RC.Config.length; i++) {
                if (RC.Config[i].type == requested_view_type) {
                    view_configuration = RC.Config[i];
                    break;
                }
            }
            var ctx = SP.ClientContext.get_current();
            var web = ctx.get_web();
            var listCollection = web.get_lists();
            var list = listCollection.getById(_spPageContextInfo.pageListId);
            var viewCollection = list.get_views();
            ctx.load(viewCollection);
            ctx.executeQueryAsync(function() {
                //get the default view
                var default_view = null;
                for (var index = 0; index < viewCollection.get_count(); index++) {
                    var current_view = viewCollection.itemAt(index);
                    if (current_view.get_defaultView()) {
                        default_view = current_view;
                        break;
                    }
                }
                for (var index = 0; index < viewCollection.get_count(); index++) {
                    var current_view = viewCollection.itemAt(index);
                    if (current_view.get_title() == view_configuration.view_name && current_view.get_personalView()) {
                        //We found a view with a same name. Cannot create the view automatically.
                        var url = current_view.get_serverRelativeUrl()
                        window.location.href = window.location.protocol + "//" + window.location.host + url + "?PageView=Personal&ShowWebPart={" + current_view.get_id() + "}";
                        return;
                    }
                }
                var default_view_fields = default_view.get_viewFields()
                ctx.load(default_view_fields);
                ctx.executeQueryAsync(function() {
                    var new_view_fields = [];

                    for (var i = 0; i < default_view_fields.get_count(); i++) {
                        new_view_fields.push(default_view_fields.getItemAtIndex(i));
                    }
                    //create a new view
                    var new_view = new SP.ViewCreationInformation();
                    new_view.set_title(view_configuration.view_name);

                    var camlQuery = new SP.CamlQuery();
                    var query = view_configuration.query;
                    camlQuery.set_viewXml(query);

                    //get view query from default view
                    //new_view.set_query(default_view.get_viewQuery());
                    new_view.set_query(camlQuery);
                    new_view.set_viewFields(new_view_fields)
                    new_view.set_paged(true);
                    new_view.set_personalView(true);
                    viewCollection.add(new_view);
                    ctx.load(viewCollection);
                    ctx.executeQueryAsync(function() {
                        //We have now created the view. To get the url of the view, be a good guy and don't infer but get the value from REST
                        RC.Helpers.GetView({ListId: _spPageContextInfo.pageListId, OData: "jQueryselect=ServerRelativeUrl", ViewName: view_configuration.view_name}).done(function(data) {
                            var url = data.d.ServerRelativeUrl
                            window.location.href = window.location.protocol + "//" + window.location.host + url;
                        }).fail(function(err) {
                            console.error(RC.Messages.ErrorRetrievingValues);
                            console.error(err);
                        })
                    }, function(sender, args) {
                        console.error(RC.Messages.ErrorViewCreation);
                        console.error(args.get_message());
                    });
                }, function(sender, args) {
                    console.error(RC.Messages.ErrorRetrievingValues);
                    console.error(args.get_message());
                })

            }, function(sender, args) {
                console.error(args.get_message());
            });
        }

        return {CreateDefaultView: CreateDefaultView, CreatePersonalView: CreatePersonalView}

    })();

    //for lack of a better word
    RC.Controller = (function() {
        function CreateView(requested_view_type) {
            RC.Helpers.CheckPermissions([SP.PermissionKind.manageLists]).done(function(hasPermission) {
                if (hasPermission) {
                    RC.Services.CreateDefaultView(requested_view_type);
                } else {
                    RC.Helpers.CheckPermissions([SP.PermissionKind.managePersonalViews]).done(function(hasPersonaViewPermission) {
                        if (hasPersonaViewPermission) {
                            RC.Services.CreatePersonalView(requested_view_type);
                        } else {
                            alert(RC.Messages.ErrorInsufficientPermission);
                        }
                    }).fail(function(err) {
                        console.error("Error retrieing permissions for user")
                        console.error(err);
                    });
                }
            }).fail(function(err) {
                console.error("Error retrieing permissions for user")
                console.error(err);
            })
        }

        function RedirectToModifyView() {
            var currentPageUrl = _spPageContextInfo.serverRequestPath;
            var path_array = currentPageUrl.split('/');
            var view_aspx = path_array[path_array.length - 1];

            var context = SP.ClientContext.get_current();
            var pagesListViews = context.get_web().get_lists().getById(_spPageContextInfo.pageListId).get_views();
            context.load(pagesListViews, 'Include(Id,ServerRelativeUrl)');
            context.executeQueryAsync(function(sender, args) {
                var viewEnumerator = pagesListViews.getEnumerator();
                while (viewEnumerator.moveNext()) {
                    var view = viewEnumerator.get_current();
                    var url = view.get_serverRelativeUrl();
                    // If url.contains(viewUrl)
                    if (url == _spPageContextInfo.serverRequestPath) {
                        var view_id = view.get_id();
                        var viewEditUrl = _spPageContextInfo.webAbsoluteUrl + "/" + _spPageContextInfo.layoutsUrl + "/ViewEdit.aspx?List=" + _spPageContextInfo.pageListId + "&View={" + view_id + "}"
                        window.location.href = viewEditUrl;
                        break;
                    }
                }
            }, function(sender, args) {
                console.error(args.get_message());
            });
        }

        return {CreateView: CreateView, RedirectToModifyView: RedirectToModifyView}
    })();

    RC.Init = function() {
        ExecuteOrDelayUntilScriptLoaded(function() {
            SP.UI.Status.removeAllStatus(true);
            //Improve this if you are handling multiple locales
            var large_list_string = RC.Constants[0].error;
            var inners = jQuery.map(jQuery.find("div[webpartid]"), function(item) {
                return item.innerHTML;
            });
            var items = inners.filter(function(item, index) {
                return item.indexOf(large_list_string) >= 0;
            });
            if (items.length > 0) {
                statusID = SP.UI.Status.addStatus("Warning:", "Your view contains too many items (> 5,000 items) and so can't be displayed. <a href='#'>More information</a>. <a href='#' onclick='javascript: RC.Controller.RedirectToModifyView()'>Modify the view</a>, or create a new one: <a href='#' onclick='javascript:RC.Controller.CreateView(0)'>Most Recent</a>, <a href='#' onclick='javascript:RC.Controller.CreateView(1)'>Last 7 days</a>, <a href='#' onclick='javascript:RC.Controller.CreateView(2)'>Last 30 days</a>.");
                SP.UI.Status.setStatusPriColor(statusID, 'yellow');
            }
        }, 'sp.js');
    }

    return RC.Init;
});
