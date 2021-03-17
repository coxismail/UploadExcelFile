        var app = angular.module('CustomerImport', []).controller("CustomerImportController", function ($scope, Transferer) {
            $scope.SelectedFileForUpload = null; //initially make it null
            $scope.ListData = null;
            $scope.ShowPermission = false;
            $scope.showLoaderi = false;
            $scope.PrepareThisFile = function (files) {
                $scope.$apply(function () {
                    $scope.Message = '';
                    $scope.SelectedFileForUpload = files[0];
                });
            }


            $scope.ParseExcel = function () {
                var ExcelTableData = new FormData();
                var file = $scope.SelectedFileForUpload;
                ExcelTableData.append('file', file);
                $scope.showLoader = true;   //show spinner
                var response = Transferer.SendExcelData(ExcelTableData);   //calling service
                response.then(function (d) {

                    $scope.ListData = d.data;
                    console.log(d.data);

                    $scope.ShowPermission = true; //showing the table after databinding
                    $scope.showLoader = false; //after success hide spinner
                    $("#tableData").DataTable({
                        "data": d.data,
                        "bDestroy":true,
                        "columns": [
                            { "data": "FirstName" },
                            { "data": "LastName" },
                            { "data": "EmailID" },
                            { "data": "City" },
                            { "data": "Country" }
                        ]
                    });
                }, function (err) {
                    console.log(err.data);
                    console.log("error in parse excel");
                });
            }


            $scope.InsertData = function () {
                $scope.showLoaderi = true;
                var response = Transferer.InsertToDB();
                response.then(function (d) {
                    var res = d.data;

                    if (res.toString().length > 0) {
                        $scope.Message = res + "  Records Inserted";
                        $scope.ShowPermission = false;   //used to disable the insert button and table after submitting data
                        angular.forEach(
                            angular.element("input[type='file']"),
                            function (inputElem) {
                                angular.element(inputElem).val(null); //used to clear the file upload after submitting data
                            });
                        $scope.SelectedFileForUpload = null;
                        $scope.showLoaderi = false;
                        //used to disable the preview button after inserting data
                    }

                }, function (err) {
                    console.log(err.data);
                    console.log("error in insertdata");
                });
            }
        });
        var Transferer =  app.service("Transferer", function ($http) {
            this.SendExcelData = function (data) {
                var request = $http({
                    method: "post",
                    withCredentials: true,
                    url: '/Home/ReadExcel',
                    data: data,
                    headers: {
                        'Content-Type': undefined
                    },
                    transformRequest: angular.identity
                });
                return request;
            }
            this.InsertToDB = function () {
                var request = $http({
                    method: "get",
                    url: '/Home/UpdateCustomerTable',
                    data: {},
                    datatype: 'json'
                });
                return request;
            }
        });

     

        
    
