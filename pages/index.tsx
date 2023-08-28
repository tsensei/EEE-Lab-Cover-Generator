import {
  Container,
  Title,
  createStyles,
  Text,
  rem,
  NumberInput,
  TextInput,
  NativeSelect,
  Button,
  px,
} from "@mantine/core";
import { DateInput, DatePickerInput } from "@mantine/dates";
import { useForm } from "@mantine/form";
import Head from "next/head";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import { saveAs } from "file-saver";
import { FormEvent, FormEventHandler } from "react";
let PizZipUtils: any = null;
if (typeof window !== "undefined") {
  import("pizzip/utils/index.js").then(function (r) {
    PizZipUtils = r;
  });
}

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

const useStyles = createStyles((theme) => ({
  rootContainer: {
    padding: "2rem 0",
    display: "flex",
    flexDirection: "column",
    minHeight: "100vh",
  },
  titleContainer: {
    textAlign: "center",
  },
  contentContainer: {
    flex: 1,
    width: "100%",
    display: "flex",
    flexDirection: "column",
    justifyContent: "center",
    padding: "2rem 1rem",
  },
  form: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: px(10),
  },
  footerContainer: {},
}));

function formatDate(date) {
  // Get the day, month, and year from the date object
  const day = date.getDate();
  const month = date.getMonth() + 1; // getMonth() returns 0-11, so we add 1
  const year = date.getFullYear();

  // Add leading zeros if day or month is a single digit
  const formattedDay = day < 10 ? `0${day}` : day;
  const formattedMonth = month < 10 ? `0${month}` : month;

  // Combine them in the DD/MM/YYYY format
  return `${formattedDay}/${formattedMonth}/${year}`;
}

const Homepage = () => {
  const { classes } = useStyles();
  const form = useForm({
    initialValues: {
      expno: 1,
      expname: "",
      oddeven: "Odd",
      groupnum: 1,
      expDate: "",
      subDate: "",
      student1: "",
      student2: "",
      student3: "",
    },
  });

  const studentData = [
    "--- Select Student Name ---",
    "Arik Islam (1)",
    "Anika Sanzida Upoma (2)",
    "Abdullah Evne Masood (3)",
    "Tahsina Sultana Afifa (4)",
    "Md. Akram Khan (5)",
    "Dipta Bhattacharjee (6)",
    "Sumaiya Rahman Soma (7)",
    "Shaila Jaman Priti (8)",
    "Aditto Raihan (9)",
    "Istiak Ahammed Rhyme (10)",
    "Md. Shakin Alam Karbo (11)",
    "Mir. Md. Ishrak Faisal (12)",
    "H.M. Mehedi Hasan (13)",
    "Md. Sarif (14)",
    "Srabon Aich (15)",
    "Ibna Afra Roza (16)",
    "Suraya Jannat Mim (17)",
    "Swapon Chandra Roy (18)",
    "Anisha Tabassum (19)",
    "Md. Ashif Mahmud Kayes (20)",
    "Abantika Paul (21)",
    "Mehedi Hasan (22)",
    "Jubayer Ahmed Sojib (23)",
    "Jobaer Hossain Tamim (24)",
    "Md. Shahria Hasan Jony (25)",
    "Sonia Akter (26)",
    "Md. Sadman Sakib (27)",
    "Dibbajothy Sarker (28)",
    "Md. Mohasin Molla (29)",
    "Md. Mahmudur Rahman Moin (30)",
    "Jotish Biswas (31)",
    "Saad Bin Ashad (32)",
    "Sharfraz Khan Hridue (33)",
    "Abdullah-Ash-Sakafy (34)",
    "Farhan Bin Rabbani (35)",
    "Md. Sadman Shihab (36)",
    "Labonya Pal (37)",
    "Ahnaf Mahbub Khan (38)",
    "Tamal Kanti Sarker (39)",
    "Nafisha Akhter (40)",
    "Md. Rushan Jamil (41)",
    "S. M. Shamiun Ferdous (42)",
    "Hasanat Ashrafy (43)",
    "Md. Nadim Mahmud Chowdhury Sizan (44)",
    "Md. Ariful Islam (45)",
    "Ahil Islam Aurnob (46)",
    "Md. Abu Bakar Siddique (47)",
    "Farhana Alam (48)",
    "Atiya Fahmida Noshin (49)",
    "Biplob Pal (50)",
    "Tasnova Shahrin (51)",
    "Most. Ishrat Jahan Mim (52)",
    "Abul Hasan Anik (53)",
    "Mst. Tasmia Sultana Sumi (54)",
    "Chowdhury Shafahid Rahman (55)",
    "Mst. Tabassum Kabir (56)",
    "N. M Rashidujjaman Masum (57)",
    "Sara Faria Sundra (58)",
    "Jubair Ahammad Akter (59)",
    "Md. Mahmudul Hasan (60)",
  ];

  const generateDocument = (e: FormEvent) => {
    e.preventDefault();

    console.log(process.env.NODE_ENV);

    loadFile(`/template.docx`, (error, content) => {
      if (error) {
        throw error;
      }

      let zip = new PizZip(content);
      let doc = new Docxtemplater().loadZip(zip);

      const {
        expDate,
        expname,
        expno,
        groupnum,
        oddeven,
        student1,
        student2,
        student3,
        subDate,
      } = form.values;

      let formateedExpDate = formatDate(expDate);
      let formattedSubDate = formatDate(subDate);

      doc.setData({
        expno: expno,
        expname: expname,
        oddeven: oddeven,
        groupnum: oddeven === "Odd" ? `A(${groupnum})` : `B(${groupnum})`,
        expDate: formateedExpDate,
        subDate: formattedSubDate,
        student1: student1,
        student2: student2,
        student3: student3,
      });

      try {
        doc.render();
      } catch (error: any) {
        const replaceErrors = (key, value) => {
          if (value instanceof Error) {
            return Object.getOwnPropertyNames(value).reduce(function (
              error,
              key
            ) {
              error[key] = value[key];
              return error;
            },
            {});
          }
          return value;
        };

        console.log(JSON.stringify({ error: error }, replaceErrors));

        if (error.properties && error.properties.errors instanceof Array) {
          const errorMessages = error.properties.errors
            .map(function (error) {
              return error.properties.explanation;
            })
            .join("\n");
          console.log("errorMessages", errorMessages);
          // errorMessages is a humanly readable message looking like this :
          // 'The tag beginning with "foobar" is unopened'
        }

        throw error;
      }

      var out = doc.getZip().generate({
        type: "blob",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      // Output the document using Data-URI
      saveAs(out, "labCover.docx");
    });
  };

  return (
    <Container className={classes.rootContainer}>
      <Head>
        <title>EEE Lab Report Generator</title>
      </Head>
      <Container className={classes.titleContainer}>
        <Title>EEE Lab Report Generator</Title>
      </Container>
      <Container className={classes.contentContainer}>
        <form onSubmit={generateDocument} className={classes.form}>
          <NumberInput
            defaultValue={1}
            placeholder="Experiment No"
            label="Experiment No"
            withAsterisk
            required
            {...form.getInputProps("expno")}
          />
          <TextInput
            placeholder="Name of the experiment"
            label="Experiment Name"
            withAsterisk
            required
            {...form.getInputProps("expname")}
          />
          <NativeSelect
            data={["Odd", "Even"]}
            label="Select your lab group"
            description="Odd or Even"
            withAsterisk
            required
            {...form.getInputProps("oddeven")}
          />
          <NumberInput
            defaultValue={1}
            placeholder="Enter your lab group number"
            label="Lab Group number"
            withAsterisk
            required
            {...form.getInputProps("groupnum")}
          />
          <DateInput
            valueFormat="DD/MM/YYYY"
            required
            label="Experiment Date"
            placeholder="Pick experiment date when lab was held"
            {...form.getInputProps("expDate")}
          />
          <DateInput
            valueFormat="DD/MM/YYYY"
            required
            label="Submission Date"
            placeholder="Pick date when this report will be submitted"
            {...form.getInputProps("subDate")}
          />
          <NativeSelect
            data={studentData}
            required
            label="Enter student 1 name"
            withAsterisk
            {...form.getInputProps("student1")}
          />
          <NativeSelect
            data={studentData}
            required
            label="Enter student 2 name"
            withAsterisk
            {...form.getInputProps("student2")}
          />
          <NativeSelect
            data={studentData}
            required
            label="Enter student 3 name"
            withAsterisk
            {...form.getInputProps("student3")}
          />
          <Button
            type="submit"
            style={{
              alignSelf: "center",
              marginTop: "1rem",
            }}
          >
            Generate cover
          </Button>
        </form>
      </Container>
      <Container className={classes.footerContainer}>
        <Text>
          Made with ❤️ by{" "}
          <Text
            component="a"
            href="https://www.facebook.com/tsenseiii/"
            target="_blank"
            variant="gradient"
            gradient={{ from: "blue", to: "cyan" }}
            inherit
          >
            tsensei
          </Text>
        </Text>
      </Container>
    </Container>
  );
};

export default Homepage;
