import React, { useEffect, useState } from "react";
import * as ExcelJS from "exceljs";
import "./App.css";

const toDataURL = (url) => {
  const promise = new Promise((resolve, reject) => {
    var xhr = new XMLHttpRequest();
    xhr.onload = function () {
      var reader = new FileReader();
      reader.readAsDataURL(xhr.response);
      reader.onloadend = function () {
        resolve({ base64Url: reader.result });
      };
    };
    xhr.open("GET", url);
    xhr.responseType = "blob";
    xhr.send();
  });

  return promise;
};

const App = () => {
  const [data, setData] = useState([]);
  useEffect(() => {
    fetch("https://dummyjson.com/products")
      .then((res) => res.json())
      .then(async (data) => {
        console.log(data);
        setData(data);
      })
      .then((json) => console.log(json));
  }, []);

  const exportExcelFile = () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("My Sheet");

    sheet.mergeCells("A1:B1");
    sheet.mergeCells("C1:D1");
    sheet.mergeCells("E1:G1");
    sheet.getCell("A2").value = "Id";
    sheet.getCell("B2").value = "Title";
    sheet.getCell("C2").value = "Brand";
    sheet.getCell("D2").value = "Type";
    sheet.getCell("E2").value = "Price";
    sheet.getCell("F2").value = "Rating";
    sheet.getCell("G2").value = "Photo";
    sheet.getRow(2).commit();
    // merge by start row, start column, end row, end column
    //sheet.properties.defaultRowHeight = 80;
    //color: { argb: "FFFF0000" }

    // for (var i = 1; i < data.products.length; i++) {
    //   sheet.properties.defaultRowHeight = 80;
    // }

    for (var i = 1; i < 8; i++) {
      sheet.getRow(2).getCell(i).border = {
        top: { style: "thick" },
        left: { style: "thick" },
        bottom: { style: "thick" },
        right: { style: "thick" },
      };

      sheet.getRow(2).getCell(i).fill = {
        fgColor: { argb: "FFFF00" },
        type: "pattern",
        pattern: "darkVertical",
      };
    }

    sheet.columns = [
      {
        header: "Id",
        key: "id",
        width: 10,
      },
      { header: "Title", key: "title", width: 32 },
      {
        header: "Brand",
        key: "brand",
        width: 20,
      },
      {
        header: "Category",
        key: "category",
        width: 20,
      },
      {
        header: "Price",
        key: "price",
        width: 15,
      },
      {
        header: "Rating",
        key: "rating",
        width: 10,
      },
      {
        header: "Photo",
        key: "thumbnail",
        width: 12,
      },
    ];

    // sheet.getColumn(3).outlineLevel = 2;
    // sheet.getRow(3).outlineLevel = 2;

    const promise = Promise.all(
      data?.products?.map(async (product, index) => {
        const rowNumber = index + 2;
        sheet.addRow({
          id: product?.id,
          title: product?.title,
          brand: product?.brand,
          category: product?.category,
          price: product?.price,
          rating: product?.rating,
        });

        let photo = product?.thumbnail;
        // "https://img.freepik.com/free-psd/rose-flower-png-isolated-transparent-background_191095-12169.jpg";
        const result = await toDataURL(photo);
        console.log(rowNumber);

        const imageId2 = workbook.addImage({
          base64: result.base64Url,
          extension: "jpeg",
        });

        sheet.addImage(imageId2, {
          tl: { col: 6, row: rowNumber },
          ext: { width: 80, height: 50 },
        });
      })
    );

    promise.then(() => {
      // const priceCol = sheet.getColumn(5);

      // priceCol.eachCell((cell) => {
      //   const cellValue = sheet.getCell(cell?.address).value;

      //   if (cellValue > 50 && cellValue < 1000) {
      //     sheet.getCell(cell?.address).fill = {
      //       type: "pattern",
      //       pattern: "solid",
      //       fgColor: { argb: "FF0000" },
      //     };
      //   }
      // });

      for (var i = 1; i < data?.products?.length + 3; i++) {
        sheet.getRow(i).height = 55;
      }

      workbook.xlsx.writeBuffer().then(function (data) {
        const blob = new Blob([data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const url = window.URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = "download.xlsx";
        anchor.click();
        window.URL.revokeObjectURL(url);
      });
    });
  };

  return (
    <>
      <div className="card">
        <button onClick={exportExcelFile}>ออกรายงาน Excel</button>
      </div>
    </>
  );
};

export default App;
