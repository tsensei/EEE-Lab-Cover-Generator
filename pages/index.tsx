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

  const generateDocument = (e: FormEvent) => {
    e.preventDefault();

    loadFile("/template.docx", (error, content) => {
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
        groupnum: groupnum,
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

      console.log("aaa");
      var out = doc.getZip().generate({
        type: "blob",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      // Output the document using Data-URI
      saveAs(out, "labCover.docx");
      console.log("aaa2");
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
          <TextInput
            required
            placeholder="Format : Name (Roll)"
            label="Enter student 1 name"
            withAsterisk
            {...form.getInputProps("student1")}
          />
          <TextInput
            required
            placeholder="Format : Name (Roll)"
            label="Enter student 2 name"
            withAsterisk
            {...form.getInputProps("student2")}
          />
          <TextInput
            required
            placeholder="Format : Name (Roll)"
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
            href="https://github.com/tsensei"
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
