package com.excel.import_export_demo.demo;

import java.util.List;

import org.springframework.stereotype.Service;

import lombok.RequiredArgsConstructor;

@Service
@RequiredArgsConstructor
public class PersonSerivce {

    private final PersonRepository personRepository;

    public List<PersonResponse> getPersons() {
        return personRepository.findAll()
                .stream()
                .map(person -> PersonResponse.builder()
                        .name(person.getName())
                        .age(person.getAge())
                        .build()
                )
                .toList();
    }

    public Long addPerson(PersonRequest personRequest) {
        return personRepository.save(Person.builder()
        .name(personRequest.name)
        .age(personRequest.age)
        .build()).getId();
    }
}
